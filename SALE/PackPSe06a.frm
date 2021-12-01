VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PackPSe06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Split A Packing Slip Item"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSe06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optInc 
      Caption         =   "Include"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      ToolTipText     =   "Include This Item"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "C&reate"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   3
      ToolTipText     =   "Create A New Packing Slip"
      Top             =   2160
      Width           =   875
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   240
      Top             =   3000
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   6735
   End
   Begin VB.ComboBox cmbItm 
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Packing Slip Item (Not Sales Order Item)"
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox cmbPsl 
      Height          =   315
      Left            =   1215
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "9"
      ToolTipText     =   "Only Qualifying Packing Slips.  Select From List"
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   4
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
      FormDesignHeight=   3570
      FormDesignWidth =   7065
   End
   Begin VB.Label lblPackSlip 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1200
      TabIndex        =   28
      ToolTipText     =   "New Packing Slip (Reserved)"
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   312
      Left            =   6600
      TabIndex        =   27
      ToolTipText     =   "Available On This Packing Slip"
      Top             =   840
      Width           =   372
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Count"
      Height          =   252
      Index           =   6
      Left            =   5760
      TabIndex        =   26
      Top             =   840
      Width           =   732
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3720
      TabIndex        =   25
      ToolTipText     =   "Part Description"
      Top             =   2880
      Width           =   3180
   End
   Begin VB.Label lblSoItem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2160
      TabIndex        =   24
      ToolTipText     =   "Sales Order Item"
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2610
      TabIndex        =   23
      ToolTipText     =   "Quantity"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblSoType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1200
      TabIndex        =   22
      ToolTipText     =   "Sales Order"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblNewItem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   315
      Left            =   2160
      TabIndex        =   20
      ToolTipText     =   "New Packing Slip Item (Reserved)"
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblNewPs 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1200
      TabIndex        =   18
      ToolTipText     =   "New Packing Slip (Reserved)"
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New PS"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblPart 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3720
      TabIndex        =   16
      ToolTipText     =   "Part Number"
      Top             =   2520
      Width           =   3180
   End
   Begin VB.Label lblSalesOrder 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1455
      TabIndex        =   15
      ToolTipText     =   "Sales Order"
      Top             =   2520
      Width           =   600
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   14
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblPrn 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3840
      TabIndex        =   13
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   11
      Top             =   1560
      Width           =   3795
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed"
      Height          =   255
      Index           =   23
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "This Function Is Intended For Packing Slips Printed, Not Shipped And Not Invoiced.  Please Select From The Lists."
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "PackPSe06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'10/6/04 new
'4/12/05 Added multiple split items
'4/29/05 Revised Lot/Inva update
'7/9/05 Reformatted Item No and lblQty (widen)
'7/28/05 Changed cmbPsl to allow numeric entry
'8/8/05 Corrected Error created by trimming the PSNUMBER
'7/27/06 Added PSPRIMARYSO to Insert
Option Explicit
'Dim rdoQry As rdoQuery
Dim cmdObj As ADODB.Command
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim iSoItem As Integer
Dim iTotActRows As Integer
Dim iTotalLots As Integer
Dim lSoNumber As Long
Dim sSoItemRev As String

Dim sLots(50, 3) As String 'Recovery of lots
Dim sActivity(150) As String 'Recovery of Inva
Dim vSplits(300, 8) As Variant
'0 = PS Item
'1 = So Number
'2 = So Item
'3 = So Item Rev
'4 = So Qty
'5 = Part Number
'6 = Include
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetNewPackslip()
'   Dim RdoGet As ADODB.Recordset
'   Dim l As Long
'   Dim n As Long
'   On Error GoTo DiaErr1
'   sSql = "SELECT CURPSNUMBER FROM ComnTable WHERE COREF=1"
'   bSqlRows = GetDataSet(RdoGet)
'   If bSqlRows Then
'      With RdoGet
'         If Len(Trim(!CURPSNUMBER)) <= 6 Then
'            l = Val(!CURPSNUMBER)
'         Else
'            l = Val(Right$(!CURPSNUMBER, 6))
'         End If
'         ClearResultSet RdoGet
'      End With
'   Else
'      l = 1
'   End If
'   lblPackSlip = Format(l + 1, "000000")
'   lblNewPs = "PS" & Format(l + 1, "000000")
'   Set RdoGet = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getpacksl"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'

   Dim ps As New ClassPackSlip
   lblPackSlip = ps.GetNextPackSlipNumber
   lblNewPs = lblPackSlip
End Sub

Private Sub GetPackslip()
   Dim RdoPsl As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PSNUMBER,PSCUST,PSDATE,PSPRINTED " _
          & "FROM PshdTable WHERE PSNUMBER='" & Trim(cmbPsl) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPsl, ES_FORWARD)
   If bSqlRows Then
      With RdoPsl
         lblCst = "" & Trim(!PSCUST)
         lblDte = Format(!PSDATE, "mm/dd/yyyy")
         lblPrn = Format(!PSPRINTED, "mm/dd/yyyy")
         ClearResultSet RdoPsl
      End With
      FindCustomer Me, lblCst
   Else
      lblCst = ""
      lblDte = ""
      lblPrn = ""
   End If
   If lblCst <> "" Then GetPSItems
   Set RdoPsl = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpackslip"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbItm_Click()
   GetSOItem
   
End Sub

Private Sub cmbItm_LostFocus()
   If Trim(cmbItm) = "" Then cmbItm = cmbItm.List(0)
   
End Sub


Private Sub cmbPsl_Click()
   GetPackslip
   
End Sub


Private Sub cmbPsl_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   If bCancel = 1 Then Exit Sub
   'cmbPsl = CheckLen(cmbPsl, 6)
   'cmbPsl = Format(Abs(Val(cmbPsl)), "000000")
   'If cmbPsl.ListCount > 0 Then
   '   If Trim(cmbPsl) = "" Then cmbPsl = cmbPsl.List(0)
   'For iList = 0 To cmbPsl.ListCount - 1
   '   If cmbPsl = cmbPsl.List(iList) Then bByte = 1
   'Next
   'If bByte = 1 Then
      GetPackslip
   'Else
   '   MsgBox "Select Or Enter The Packing Slip Number From One On The List.", _
   '      vbInformation, Caption
   '   If cmbPsl.ListCount > 0 Then cmbPsl = cmbPsl.List(0)
   'End If
   
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
      OpenHelpContext 2206
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSplit_Click()
   Dim bResponse As Byte
   Dim iList As Integer
   Dim sMsg As String
   
   For iList = 0 To cmbItm.ListCount - 1
      bResponse = bResponse + vSplits(iList, 6)
   Next
   If bResponse = 0 Then
      MsgBox "No Items Are Selected To Be Split.", _
         vbInformation, Caption
      Exit Sub
   End If
   If lSoNumber = 0 Then
      MsgBox "Requires A Valid Sales Order.", _
         vbInformation, Caption
   Else
      sMsg = "This Function Removes The Selected Item(s)" & vbCrLf _
             & "From " & cmbPsl & " And Creates A New Packing" & vbCrLf _
             & "Packing Slip " & lblNewPs & ". Continue?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then CreateSplit Else CancelTrans
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT PIPACKSLIP,PIITNO,PIQTY,PIPART,PISONUMBER,PISOITEM," _
          & "PISOREV,PARTREF,PARTNUM,PADESC,SONUMBER,SOTYPE From PsitTable," _
          & "PartTable,SohdTable WHERE (PIPART=PARTREF AND PISONUMBER=SONUMBER) " _
          & "AND PIPACKSLIP= ? AND PIITNO= ? "
   'Set rdoQry = RdoCon.CreateQuery("", sSql)
   'rdoQry.MaxRows = 1
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql
   
   Dim prmObj As ADODB.Parameter
   Dim prmObj1 As ADODB.Parameter
   
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adChar
   prmObj.Size = 8
   cmdObj.Parameters.Append prmObj
   
   Set prmObj1 = New ADODB.Parameter
   prmObj1.Type = adInteger
   cmdObj.Parameters.Append prmObj1
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set cmdObj = Nothing
   FormUnload
   Set PackPSe06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   Dim RdoPsl As ADODB.Recordset
   On Error GoTo DiaErr1
   cmbPsl.Clear
   cmbItm.Clear
   GetNewPackslip
   Timer1.Enabled = True
'   sSql = "SELECT DISTINCT PSNUMBER,PIPACKSLIP FROM " _
'          & "PshdTable,PsitTable WHERE (PSPRINTED IS NOT NULL AND " _
'          & "PSINVOICE=0 AND PSSHIPPED=0 AND PSNUMBER=PIPACKSLIP) "
   
   'join with PoitTable to exclude empty packing slips
   sSql = "SELECT DISTINCT PSNUMBER" & vbCrLf _
      & "FROM PshdTable" & vbCrLf _
      & "JOIN PsitTable on PSNUMBER = PIPACKSLIP" & vbCrLf _
      & "WHERE PSPRINTED IS NOT NULL" & vbCrLf _
      & "AND PSINVOICE=0 AND PSSHIPPED=0" & vbCrLf _
      & "ORDER BY PSNUMBER"
   LoadComboBoxAndSelect cmbPsl
   
'   bSqlRows = GetDataSet(RdoPsl, ES_FORWARD)
'   If bSqlRows Then
'      With RdoPsl
'         Do Until .EOF
'            AddComboStr cmbPsl.hWnd, "" & Trim(Right$(.Fields(0), 6))
'            .MoveNext
'         Loop
'         ClearResultSet RdoPsl
'      End With
'   End If
   If cmbPsl.ListCount > 0 Then
      GetPackslip
   Else
      MsgBox "No Qualifying Packing Slips Found.", _
         vbInformation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetPSItems()
   Dim RdoItm As ADODB.Recordset
   Dim iList As Integer
   cmbItm.Clear
   On Error GoTo DiaErr1
   Erase vSplits
   iList = -1
   optInc.Enabled = False
   sSql = "SELECT PIPACKSLIP,PIITNO FROM PsitTable WHERE " _
          & "PIPACKSLIP='" & cmbPsl & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   If bSqlRows Then
      With RdoItm
         Do Until .EOF
            iList = iList + 1
            vSplits(iList, 0) = !PIITNO
            cmbItm.AddItem "" & Trim$(str$(!PIITNO))
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
   End If
   Set RdoItm = Nothing
   If cmbItm.ListCount > 0 Then
      optInc.Enabled = True
      If cmbItm.ListIndex < 0 Then cmbItm.ListIndex = 0
      cmbItm = cmbItm.List(0)
      cmdSplit.Enabled = True
      GetSOItem
   Else
      cmdSplit.Enabled = False
      lblSalesOrder = ""
      lblPart = ""
   End If
   Set RdoItm = Nothing
   lblCount = cmbItm.ListCount
   Exit Sub
   
DiaErr1:
   sProcName = "getpsitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetSOItem()
   Dim RdoSit As ADODB.Recordset
   Dim iList As Integer
   On Error GoTo DiaErr1
   If cmbItm.ListIndex < 0 Then iList = 0 _
                                        Else iList = cmbItm.ListIndex
'   rdoQry(0) = cmbPsl
'   rdoQry(1) = Val(cmbItm)
'   bSqlRows = GetQuerySet(RdoSit, rdoQry, ES_FORWARD)
   cmdObj.Parameters(0).Value = cmbPsl
   cmdObj.Parameters(1).Value = Val(cmbItm)
   bSqlRows = clsADOCon.GetQuerySet(RdoSit, cmdObj, ES_FORWARD, True)
      
   If bSqlRows Then
      With RdoSit
         lblSoType = "" & Trim(!SOTYPE)
         lblSalesOrder = Format(!SoNumber, SO_NUM_FORMAT)
         lblSoItem = "" & Trim$(str$(!PISOITEM)) & Trim(!PISOREV)
         lblQty = Format(!PIQTY, ES_QuantityDataFormat)
         lblPart = "" & Trim(!PartNum)
         lblDesc = "" & Trim(!PADESC)
         iSoItem = !PISOITEM
         lSoNumber = !SoNumber
         sSoItemRev = "" & Trim(!PISOREV)
         vSplits(iList, 1) = !SoNumber
         vSplits(iList, 2) = !PISOITEM
         vSplits(iList, 3) = !PISOREV
         vSplits(iList, 4) = !PIQTY
         vSplits(iList, 5) = Compress(!PartNum)
         optInc.Value = Val(vSplits(iList, 6))
         ClearResultSet RdoSit
      End With
   Else
      lblSoType = ""
      lblSalesOrder = ""
      lblSoItem = ""
      lblQty = ""
      lblPart = ""
      lblDesc = ""
      iSoItem = 0
      lSoNumber = 0
      sSoItemRev = ""
   End If
   Set RdoSit = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub




Private Sub lblNewPs_Change()
   If UCase$(lblNewPs) = "ERROR" Then lblNewPs.ForeColor = ES_RED _
             Else lblNewPs.ForeColor = vbBlack
   
End Sub

Private Sub optInc_Click()
   If optInc.Value = vbChecked Then cmbPsl.Enabled = False
   If cmbItm.ListIndex < 0 Then cmbItm.ListIndex = 0
   vSplits(cmbItm.ListIndex, 6) = optInc.Value
   
End Sub

Private Sub Timer1_Timer()
   GetNewPackslip
   
End Sub



Private Sub CreateSplit()
   Dim RdoNew As ADODB.Recordset
   Dim iRow As Integer
   Dim iList As Integer
   Dim iPsItm As Integer
   Dim vAdate As Variant
   Dim oldps As String, newps As String
   oldps = cmbPsl
   newps = lblNewPs
   
   Timer1.Enabled = False
   vAdate = Format(GetServerDateTime(), "mm/dd/yyyy hh:mm")
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   'On Error Resume Next
   'new Packing slip
   sSql = "SELECT * FROM PshdTable WHERE PSNUMBER='" & oldps & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNew, ES_STATIC)
   If bSqlRows Then
      With RdoNew
         sSql = "INSERT INTO PshdTable (PSNUMBER,PSTYPE,PSCUST,PSDATE,PSPRINTED," _
                & "PSVIA,PSTERMS,PSSTNAME,PSSTADR,PSSHIPPRINT,PSSHIPPED,PSPRIMARYSO) " _
                & "VALUES ('" & newps & "',1,'" & Trim(!PSCUST) & "','" _
                & vAdate & "','" & vAdate & "','" & Trim(!PSVIA) & "','" _
                & Trim(!PSTERMS) & "','" & Trim(!PSSTNAME) & "','" _
                & Trim(!PSSTADR) & "',1,0," & Val(lblSalesOrder) & ")"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         ClearResultSet RdoNew
      End With
   End If
   
   Set RdoNew = Nothing
   
'   sSql = "UPDATE ComnTable SET CURPSNUMBER='" & lblPackSlip & "' WHERE COREF=1"
'   RdoCon.Execute sSql, rdExecDirect
   Dim ps As New ClassPackSlip
   ps.SaveLastPSNumber newps
   
   'Update PS rows
   For iList = 0 To cmbItm.ListCount - 1
      If Val(vSplits(iList, 6)) = 1 Then
         iPsItm = iPsItm + 1
         sSql = "UPDATE PsitTable SET PIPACKSLIP='" & newps & "'," _
                & "PIITNO=" & iPsItm & " WHERE (PIPACKSLIP='" & oldps & "' AND " _
                & "PIITNO=" & Val(vSplits(iList, 0)) & ")"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         
         'Update SoitTable @@@
         sSql = "UPDATE SoitTable SET ITPSNUMBER='" & newps & "'," _
                & "ITPSITEM=" & iPsItm & " WHERE (ITSO=" & Val(vSplits(iList, 1)) & " AND ITNUMBER=" _
                & Val(vSplits(iList, 2)) & " AND ITREV='" & vSplits(iList, 3) & "')"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         
         'Update LoitTable
         sSql = "UPDATE LoitTable SET LOIPSNUMBER='" & newps & "'," _
                & "LOIPSITEM=" & iPsItm & " WHERE (LOIPSNUMBER='" & oldps & "' " _
                & "AND LOIPSITEM=" & vSplits(iList, 0) & ")"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         
         sSql = "UPDATE InvaTable SET INPSNUMBER='" & newps _
                & "',INPSITEM=" & iPsItm & ",INREF2='" & newps & "-" & iPsItm & "' " _
                & "WHERE (INPSNUMBER='" & oldps & "' AND INPSITEM=" & Val(vSplits(iList, 0)) & ")"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
      End If
   Next
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      SysMsg "The Split Was Completed.", True
      lblSoType = ""
      lblSalesOrder = ""
      lblSoItem = ""
      lblQty = ""
      lblPart = ""
      lblDesc = ""
      iSoItem = 0
      lSoNumber = 0
      sSoItemRev = ""
      FillCombo
   Else
      clsADOCon.RollbackTrans
      MsgBox Err.Description & vbCrLf _
         & "Couldn't Successfully Complete The Transaction.", _
         vbInformation, Caption
   End If
   
End Sub

Public Sub GetLots(PsItemNo As Integer)
   '    Dim RdoLots As ADODB.Recordset
   '    Erase sLots
   '    iTotalLots = 0
   '    On Error GoTo DiaErr1
   '    sSql = "SELECT LOINUMBER,LOIRECORD,LOIACTIVITY FROM LoitTable " _
   '        & "WHERE (LOIPSNUMBER='PS" & Trim(cmbPsl) & "' AND LOIPSITEM=" _
   '        & PsItemNo & " AND LOITYPE=25 AND LOIADATE>='" & lblPrn & " 00:00') "
   '    bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
   '        If bSqlRows Then
   '            With RdoLots
   '                Do Until .EOF
   '                    iTotalLots = iTotalLots + 1
   '                    sLots(iTotalLots, 0) = "" & Trim(!LOINUMBER)
   '                    sLots(iTotalLots, 1) = Trim$(Str$(!LOIRECORD))
   '                    sLots(iTotalLots, 2) = Trim$(Str$(!LOIACTIVITY))
   '                    .MoveNext
   '                Loop
   '                .Cancel
   '            End With
   '        End If
   '    Set RdoLots = Nothing
   '    Exit Sub
   '
   'DiaErr1:
   '    iTotalLots = 0
   
End Sub

Private Sub GetActivity()
   '    Dim rdoAct As ADODB.Recordset
   '    Erase sActivity
   '    iTotActRows = 0
   '    sSql = "SELECT INTYPE, MAX(INNUMBER) AS ACTRECORD," _
   '        & "INLOTNUMBER FROM InvaTable WHERE (INTYPE=25 AND INPSNUMBER='PS" & Trim(cmbPsl) _
   '        & "' AND INPSITEM=" & Val(cmbItm) & ") group by INTYPE,INLOTNUMBER"
   '    bSqlRows = GetDataSet(rdoAct, ES_FORWARD)
   '        If bSqlRows Then
   '            With rdoAct
   '                Do Until .EOF
   '                    iTotActRows = iTotActRows + 1
   '                    sActivity(iTotActRows) = "" & Trim(!ACTRECORD)
   '                    .MoveNext
   '                Loop
   '                .Cancel
   '            End With
   '        End If
   '    Set rdoAct = Nothing
   '    Exit Sub
   '
   'DiaErr1:
   '    iTotActRows = 0
End Sub
