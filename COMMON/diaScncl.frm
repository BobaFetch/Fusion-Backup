VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaScncl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Close A Manufacturing Order"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optExp 
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3120
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox optUnpicked 
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      ToolTipText     =   "Workstation Setting - Allow To Close With Unpicked Items"
      Top             =   2880
      Width           =   495
   End
   Begin VB.CheckBox optInv 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   22
      Top             =   320
      Width           =   495
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "diaScncl.frx":0000
      Height          =   350
      Left            =   6120
      Picture         =   "diaScncl.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "View Last Closed Run Log (Requires A Text Viewer) "
      Top             =   1920
      Width           =   360
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1575
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Qualified Part Numbers (CO)"
      Top             =   720
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Contains Qualified Runs"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "M&O Close"
      Height          =   315
      Left            =   5640
      TabIndex        =   3
      ToolTipText     =   "Press To Close this Manufacturing Order"
      Top             =   1560
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5640
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaScncl.frx":09B4
      PictureDn       =   "diaScncl.frx":0AFA
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   3360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3600
      FormDesignWidth =   6600
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore Part Type 5's (Expendables)"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   9
      Left            =   240
      TabIndex        =   26
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore Unpicked Items"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   8
      Left            =   240
      TabIndex        =   25
      ToolTipText     =   "Workstation Setting - Allow To Close With Unpicked Items"
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Verify Invoicing (PO) Before Closing MO"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   74
      Left            =   240
      TabIndex        =   21
      ToolTipText     =   "Test Allocated PO Items For Invoices (System Setting)"
      Top             =   320
      Width           =   3495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   255
      Index           =   7
      Left            =   5040
      TabIndex        =   19
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   17
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   16
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblCode 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   15
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblLvl 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6000
      TabIndex        =   14
      ToolTipText     =   "Part Type"
      Top             =   720
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Closed"
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   13
      Top             =   1575
      Width           =   1095
   End
   Begin VB.Label lblDte 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Completed"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblStat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "diaScncl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/26/00 Accounts
'1/2/03 Returned the code to trap the Close/Completion Dates
'3/18/03 Added updating LohdTable (lot costs)
'11/11/04 All areas revised and Log added
'
'*** 11/11/04 EMail to Larry/Nathan telling them to check the code and test it
'*** 11/29/04 Telecon Larry.  He has ignored the Email. Re-iterated the necessity
'*** 12/16/04 Tested at JEVCO and it is okay except won't close some
'12/17/04 Reset bCantClose flag for ensuing MO's
'1/6/05 Changed erroneous references to cmbPrt/cmbRun
'*** 1/6/05 apparently it has not been tested yet (see above)
'2/22/05 Reacted to a fax from JEVCO as a result of telecon with Larry
'     It is obvious that Larry hasn't tested the function.
'3/15/05 Added Unpicked switch (AWI)
'5/24/05 Added option to ignore Part Type 5 parts
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bCantClose As Byte
Dim bOnLoad As Byte
Dim bGoodPrt As Byte
Dim bGoodRun As Byte
Dim bLotsOn As Byte
Dim iLogFile As Integer
Dim iLogRecord As Integer
Dim iTotalPicks As Integer
Dim lRunno As Long
Dim sPartNumber As String

Dim cYield As Currency
Dim cRunExp As Currency
Dim cRunHours As Currency
Dim cRunLabor As Currency
Dim cRunMatl As Currency
Dim cRunOvHd As Currency
Dim cStdCost As Currency

Dim sLotNumber As String
Dim sPartLots(100, 4) As String '0=Part Number
'1=Lots 0/1
'2=Standard Cost
'3=Quantity Picked

'WIP
Dim sInvLabAcct As String
Dim sInvMatAcct As String
Dim sInvExpAcct As String
Dim sInvOhdAcct As String
Dim sCgsAcct As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function OpenLog(Optional LeaveOpen As Boolean) As Integer
   On Error GoTo DiaErr1
   OpenLog = FreeFile
   Open sFilePath & "CloseRuns.txt" For Output Lock Read Write _
      As OpenLog
   Print #OpenLog, "Close Runs " & Format(ES_SYSDATE, "mm/dd/yy") & " (Dumps Log After Each Closed Run)"
   If optInv.Value = vbChecked Then
      Print #OpenLog, "Invoice Checking is Turned On"
   Else
      Print #OpenLog, "Invoice Checking is Turned Off"
   End If
   Print #OpenLog, "Requires Attention *"
   If Not LeaveOpen Then Close #OpenLog
   Exit Function
   
DiaErr1:
   OpenLog = 0
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   MouseCursor 0
   cmdCan.enabled = True
   If Err = 55 Then
      cmdDel.enabled = False
      MsgBox "Another Process Is Creating An Using This Function." & vbCr _
         & "Cannot Process Your Request.", _
         vbExclamation, Caption
      On Error GoTo 0
   Else
      DoModuleErrors Me
   End If
   
End Function

Private Sub GetWipAccounts()
   Dim b As Byte
   sProcName = "getlaboracct"
   sInvLabAcct = GetLaborAcct(sPartNumber, lblCode, Val(lblLvl))
   sProcName = "getexpenseacct"
   sInvExpAcct = GetExpenseAcct(sPartNumber, lblCode, Val(lblLvl))
   sProcName = "getmaterialacct"
   sInvMatAcct = GetMaterialAcct(sPartNumber, lblCode, Val(lblLvl))
   sProcName = "getoverheadacct"
   sInvOhdAcct = GetOverHeadAcct(sPartNumber, lblCode, Val(lblLvl))
   sProcName = "getcgsaccount"
   b = GetCGSAccounts(sCgsAcct)
   
End Sub

Private Function GetCGSAccounts(CostOfGoods As String) As Byte
   Dim rdoAct As ADODB.Recordset
   Dim bType As Byte
   Dim sPcode As String
   
   'Use current Part
   bType = Val(lblLvl)
   sSql = "SELECT PAPRODCODE,PACGSMATACCT FROM PartTable WHERE PARTREF='" & sPartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         sPcode = "" & Trim(!PAPRODCODE)
         CostOfGoods = "" & Trim(!PACGSMATACCT)
         .Cancel
      End With
   End If
   If CostOfGoods = "" Then
      'None in one or any, try Product code
      sSql = "SELECT PCCGSMATACCT FROM PcodTable WHERE PCREF='" _
             & sPcode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            If CostOfGoods = "" Then CostOfGoods = "" & Trim(!PCCGSMATACCT)
            .Cancel
         End With
      End If
   End If
   If CostOfGoods = "" Then
      'Still none, we'll check the common
      sSql = "SELECT COCGSMATACCT" & Trim(str(bType)) & " " _
             & "FROM ComnTable WHERE COREF=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            If CostOfGoods = "" Then CostOfGoods = "" & Trim(.Fields(0))
            .Cancel
         End With
      End If
   End If
   'After this excercise, we'll give up if none are found
   Set rdoAct = Nothing
   
End Function

Private Sub GetMaterialCosts()
   Dim RdoMat As ADODB.Recordset
   Dim iList As Integer
   Dim cLotCost As Currency
   Dim cQuantity As Currency
   Dim cStdCost As Currency
   Dim sFlag As String
   Dim sLotFlag As String
   
   cRunMatl = 0
   If bLotsOn Then sLotFlag = "From Lots Or Standard Cost:" _
                              Else sLotFlag = "From Standard Cost (Lots Off):"
   sProcName = "getmatlcosts"
   iLogRecord = iLogRecord + 1
   Print #iLogFile, vbCr
   iLogRecord = iLogRecord + 1
   sFlag = ""
   Print #iLogFile, "Material Costs " & sLotFlag
   If iTotalPicks > 0 Then
      For iList = 1 To iTotalPicks
         cQuantity = Val(sPartLots(iList, 3))
         If cQuantity > 0 Then
            cLotCost = 0
            If Val(sPartLots(iList, 1)) = 1 Then
               'lots - Get them
               sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTPARTREF,LOTUNITCOST," _
                      & "LOINUMBER,LOITYPE, LOIQUANTITY, LOIMOPARTREF, LOIMORUNNO " _
                      & "FROM LohdTable,LoitTable WHERE (LOTNUMBER=LOINUMBER AND " _
                      & "LOTPARTREF='" & Compress(sPartLots(iList, 0)) & "' AND " _
                      & "LOITYPE=10) AND (LOIMOPARTREF='" & Compress(cmbPrt) & "' " _
                      & "AND LOIMORUNNO=" & Val(cmbRun) & ")"
               bSqlRows = clsADOCon.GetDataSet(sSql, RdoMat, ES_FORWARD)
               If bSqlRows Then
                  With RdoMat
                     Do Until .EOF
                        iLogRecord = iLogRecord + 1
                        If !LotUnitCost > 0 Then
                           'cLotCost = Format(!LOTUNITCOST, "#####0.000")
                           cLotCost = !LotUnitCost
                           cQuantity = Abs(!LOIQUANTITY)
                           Print #iLogFile, sPartLots(iList, 0) & " Is Lot Tracked And Was Costed At Lot Unit Cost (Total): "; Format$(cQuantity * cLotCost, "#####0.000")
                           cRunMatl = cRunMatl + (cQuantity * cLotCost)
                        Else
                           cLotCost = 0
                           bCantClose = 1
                           sFlag = "* "
                           Print #iLogFile, sFlag & sPartLots(iList, 0) & " Contains An Uncosted Lot " & Trim(!LOTUSERLOTID)
                        End If
                        .MoveNext
                     Loop
                     .Cancel
                  End With
               End If
            Else
               'No Lots-Standard cost
               cStdCost = Val(sPartLots(iList, 2))
               iLogRecord = iLogRecord + 1
               If cStdCost = 0 Then
                  sFlag = "* "
                  bCantClose = 1
                  Print #iLogFile, sFlag & sPartLots(iList, 0) & " Is Not Lot Tracked And Has No Standard Cost"
               Else
                  cRunMatl = cRunMatl + (cQuantity * cStdCost)
                  Print #iLogFile, sPartLots(iList, 0) & " Is Not Lot Tracked And Was Costed At Standard (Total): "; Format$(cQuantity * cStdCost, "#####0.000")
               End If
            End If
         Else
            'Previously Reported unpicked
         End If
      Next
   Else
      iLogRecord = iLogRecord + 1
      Print #iLogFile, "No Material Costs Recorded "
   End If
   Set RdoMat = Nothing
   
End Sub

Private Sub GetLaborCosts()
   Dim RdoLab As ADODB.Recordset
   cRunHours = 0
   cRunOvHd = 0
   cRunLabor = 0
   
   sProcName = "getlaborcosts"
   iLogRecord = iLogRecord + 1
   Print #iLogFile, vbCr
   iLogRecord = iLogRecord + 1
   Print #iLogFile, "Labor Costs:"
   sSql = "SELECT TCCARD,TCHOURS,TCTIME,TCRATE,TCOHRATE," _
          & "TCPARTREF,TCRUNNO,TMDATE FROM TcitTable,TchdTable WHERE " _
          & "(TCPARTREF='" & sPartNumber & "' AND TCRUNNO=" _
          & lRunno & ") AND TCCARD=TMCARD ORDER BY TCCARD"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLab, ES_FORWARD)
   If bSqlRows Then
      With RdoLab
         Do Until .EOF
            cRunHours = cRunHours + !TCHOURS
            cRunOvHd = cRunOvHd + (!TCOHRATE * !TCHOURS)
            cRunLabor = cRunLabor + (!TCRATE * !TCHOURS)
            .MoveNext
         Loop
         .Cancel
      End With
      iLogRecord = iLogRecord + 1
      Print #iLogFile, "Hours: " & Format$(cRunHours, "#####0.00"), _
         " Labor: "; Format$(cRunLabor, "#####0.00"), _
         " Overhead: "; Format$(cRunOvHd, "#####0.00")
   Else
      iLogRecord = iLogRecord + 1
      Print #iLogFile, "Hours: " & Format$(cRunHours, "#####0.00"), _
         " Labor: "; Format$(cRunLabor, "#####0.00"), _
         " Overhead: "; Format$(cRunOvHd, "#####0.00")
      iLogRecord = iLogRecord + 1
      Print #iLogFile, "No Labor Costs Recorded"
   End If
   Set RdoLab = Nothing
   
End Sub

Private Function GetExpenseCosts() As Byte
   Dim RdoExp As ADODB.Recordset
   Dim cMOEXPENSE As Currency
   Dim cFREIGHT As Currency
   Dim cTAXES As Currency
   Dim cNOPO As Currency
   cRunExp = 0
   
   sProcName = "getexpensecos"
   
   'Purchased Expense Items
   iLogRecord = iLogRecord + 1
   Print #iLogFile, vbCr
   iLogRecord = iLogRecord + 1
   Print #iLogFile, "Invoiced Purchase Order Expenses (Services):"
   sSql = "SELECT PINUMBER,PIPART,PIITEM,PIREV,PITYPE,PIAQTY,PIAMT," _
          & "PARTREF,PARTNUM FROM PoitTable,PartTable WHERE (PIRUNPART='" _
          & Compress(cmbPrt) & "' AND PIRUNNO=" & Val(cmbRun) & " AND " _
          & "PITYPE=17 AND PALEVEL=7) AND PIPART=PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoExp, ES_FORWARD)
   If bSqlRows Then
      With RdoExp
         Do Until .EOF
            cMOEXPENSE = cMOEXPENSE + (!PIAQTY * !PIAMT)
            iLogRecord = iLogRecord + 1
            Print #iLogFile, Format(!PINUMBER, "00000"), !PIITEM & !PIREV, _
                                    !PartNum, " Qty: " & Format(!PIAQTY, "#####0.000"), _
                                    " Cost:" & Format(!PIAMT, "#####0.000")
            .MoveNext
         Loop
         .Cancel
      End With
   Else
      iLogRecord = iLogRecord + 1
      Print #iLogFile, "No Invoiced Purchase Order Expenses Items"
   End If
   'Tax and freight
   iLogRecord = iLogRecord + 1
   Print #iLogFile, vbCr
   iLogRecord = iLogRecord + 1
   Print #iLogFile, "Invoiced Purchase Order Freight And Tax:"
   sSql = "SELECT SUM(VIFREIGHT) AS FREIGHT,SUM(VITAX) AS TAX FROM VihdTable," _
          & "ViitTable WHERE VINO=VITNO AND (VITMO='" _
          & sPartNumber & "' AND VITMORUN=" & lRunno & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoExp, ES_FORWARD)
   If bSqlRows Then
      With RdoExp
         If Not IsNull(!FREIGHT) Then
            cFREIGHT = cFREIGHT + !FREIGHT
         End If
         If Not IsNull(!tax) Then
            cTAXES = cTAXES + !tax
         End If
         .Cancel
      End With
      iLogRecord = iLogRecord + 1
      Print #iLogFile, "Total Freight: " & Format(cFREIGHT, "#####0.000"), _
         "Total Taxes: " & Format(cTAXES, "#####0.000")
   Else
      iLogRecord = iLogRecord + 1
      Print #iLogFile, "Total Freight: " & Format(cFREIGHT, "#####0.000"), _
         "Total Taxes: " & Format(cTAXES, "#####0.000")
   End If
   
   'Invoices without PO's
   iLogRecord = iLogRecord + 1
   Print #iLogFile, vbCr
   iLogRecord = iLogRecord + 1
   Print #iLogFile, "Invoiced Without A Purchase Order:"
   sSql = "SELECT SUM(VITQTY*VITCOST) AS SUMCOST FROM ViitTable WHERE " _
          & "(VITPO=0 AND VITPOITEM=0) AND (VITMO='" & Compress(cmbPrt) _
          & "' AND VITMORUN=" & Val(cmbRun) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoExp, ES_FORWARD)
   If bSqlRows Then
      With RdoExp
         If Not IsNull(!SUMCOST) Then _
                       cNOPO = cNOPO + !SUMCOST Else cNOPO = 0
         .Cancel
         iLogRecord = iLogRecord + 1
         Print #iLogFile, "Total Invoiced Without PO Items " & Format$(cNOPO, "#####0.000")
      End With
   Else
      iLogRecord = iLogRecord + 1
      Print #iLogFile, "No Invoiced Without PO Items Found"
   End If
   cRunExp = cMOEXPENSE + cFREIGHT + cTAXES + cNOPO
   Set RdoExp = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getexpensec"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbPrt_Click()
   bGoodPrt = GetRunPart()
   GetRuns
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) Then
      bGoodPrt = GetRunPart()
      GetRuns
   End If
   
End Sub


Private Sub cmbRun_Click()
   bGoodRun = GetCurrRun()
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   bGoodRun = GetCurrRun()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   cmbPrt = ""
   
End Sub


Private Sub cmdDel_Click()
   Dim RdoQty As ADODB.Recordset
   Dim bByte As Byte
   Dim lClose As Long
   Dim lComplete As Long
   
   bCantClose = 0
   sJournalID = GetOpenJournal("IJ", Format$(txtDte, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      bByte = 1
   Else
      If sJournalID = "" Then bByte = 0 Else bByte = 1
   End If
   If bByte = 0 Then
      MsgBox "There Is No Open Inventory Journal For The Period " & txtDte & ".", _
         vbInformation, Caption
      Exit Sub
   End If
   bGoodRun = GetCurrRun()
   If bGoodRun = 1 Then
      lClose = DateValue(Format(txtDte, "yyyy,mm,dd"))
      lComplete = DateValue(Format(lblDte, "yyyy,mm,dd"))
      If lClose < lComplete Then
         MsgBox "The Date of Closure Cannot Be Before The" & vbCr _
            & "Completion Date.", _
            vbInformation, Caption
         Exit Sub
      End If
      sSql = "DELETE FROM EsReportClosedRuns WHERE CR_MONUMBER='" _
             & Trim(cmbPrt) & "' AND CR_RUN=" & Val(cmbRun) & " "
      clsADOCon.ExecuteSQL sSql
      
      iLogFile = OpenLog(True)
      iLogRecord = 1
      Print #iLogFile, vbCr
      iLogRecord = iLogRecord + 1
      Print #iLogFile, "Close MO " & cmbPrt & " Run " & cmbRun
      GetUnInvoicedPoItems '
      GetPickList '
      GetExpenseCosts '
      GetLaborCosts '
      GetMaterialCosts '
      If bCantClose = 0 Then
         CloseMO
      Else
         iLogRecord = iLogRecord + 1
         Print #iLogFile, vbCr
         iLogRecord = iLogRecord + 1
         Print #iLogFile, "Manufacturing Order " & cmbPrt & " Run " & cmbRun & " Was Not Closed"
         CloseLog
         On Error Resume Next
         MsgBox "Cannot Close This MO Run. See Log.", _
            vbInformation, Caption
         bByte = MsgBox("Would You Like To View The Log Now?", ES_YESQUESTION, Caption)
         If bByte = vbYes Then OpenWebPage sFilePath & "CloseRuns.txt"
      End If
   Else
      MsgBox "You Must Select A Valid Run.", _
         vbInformation, Caption
   End If
   Set RdoQty = Nothing
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "4153"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdVew_Click()
   On Error GoTo DiaErr1
   If Dir(sFilePath & "CloseRuns.txt") <> "" Then
      OpenWebPage sFilePath & "CloseRuns.txt"
   Else
      MsgBox "There Is No Current Close Runs Log.", vbInformation, Caption
   End If
   Exit Sub
   
DiaErr1:
   MsgBox "Either There Is No Current Log Or " & vbCr _
      & "CloseRuns.txt In " & sFilePath & " " & vbCr, _
      vbInformation, Caption
   On Error GoTo 0
   
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CheckInvoicing
      b = CheckInvJournal()
      bLotsOn = CheckLotStatus
      iLogFile = OpenLog()
      If b = 1 Then FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetSettings
   sSql = "SELECT RUNREF,RUNNO FROM RunsTable WHERE " _
          & "RUNREF= ? AND RUNSTATUS='CO' "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   AdoQry.parameters.Append AdoParameter
   
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveSettings
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   CloseLog
   FormUnload
   Set diaScncl = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbPrt.Clear
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF FROM " _
          & "RunsTable,PartTable WHERE PARTREF=RUNREF AND " _
          & "RUNSTATUS='CO' ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      bGoodPrt = GetRunPart()
      GetRuns
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetRuns()
   Dim RdoRns As ADODB.Recordset
   cmbRun.Clear
   sPartNumber = Compress(cmbPrt)
   AdoQry.parameters(0).Value = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry, ES_FORWARD)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoRns = Nothing
   If cmbRun.ListCount > 0 Then
      cmbRun = cmbRun.List(0)
      bGoodRun = GetCurrRun()
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetCurrRun() As Byte
   Dim RdoRun As ADODB.Recordset
   
   lRunno = Val(cmbRun)
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,RUNYIELD,RUNCOMPLETE,RUNLOTNUMBER FROM RunsTable " _
          & "WHERE RUNREF='" & Compress(cmbPrt) & "' AND RUNNO=" & lRunno & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         lblStat = "" & Trim(!RUNSTATUS)
         lblDte = Format(!RUNCOMPLETE, "mm/dd/yy")
         sLotNumber = "" & Trim(!RUNLOTNUMBER)
         lblQty = Format(!RUNYIELD, ES_QuantityDataFormat)
         cYield = !RUNYIELD
         .Cancel
      End With
   Else
      cYield = 0
      lblStat = "**"
      lblDte = ""
      sLotNumber = ""
   End If
   If lblStat = "CO" Then
      GetCurrRun = 1
   Else
      GetCurrRun = 0
      lblDte = ""
      lblStat = "**"
      sLotNumber = ""
   End If
   Set RdoRun = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcurrrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc_Change()
   If Left(lblDsc, 10) = "*** Part N" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub lblStat_Change()
   If lblStat = "**" Then
      lblStat.ForeColor = ES_RED
   Else
      lblStat.ForeColor = Es_TextForeColor
   End If
   
End Sub





Private Sub CloseMO()
   Dim RdoInv As ADODB.Recordset
   Dim bResponse As Byte
   Dim lRunno As Long
   Dim lInRecord As Long
   Dim cRunCost As Currency
   Dim sMsg As String
   Dim sPart As String
   Dim vAdate As Variant
   
   vAdate = GetServerDateTime()
   sPart = Compress(cmbPrt)
   lRunno = Val(cmbRun)
   On Error GoTo DiaErr1
   sMsg = "This Closes The MO To All Functions." & vbCrLf _
          & "Do You Really Want To Close This MO?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      cRunCost = cRunOvHd + cRunMatl + cRunExp + cRunLabor
      If cYield = 0 Then cYield = 1
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sSql = "UPDATE RunsTable SET RUNSTATUS='CL'," _
             & "RUNCOST=" & cRunCost & "," _
             & "RUNOHCOST=" & cRunOvHd & "," _
             & "RUNCMATL=" & cRunMatl & "," _
             & "RUNCEXP=" & cRunExp & "," _
             & "RUNCHRS=" & cRunHours & "," _
             & "RUNCLAB=" & cRunLabor & "," _
             & "RUNCLOSED='" & txtDte & "'," _
             & "RUNREVBY='" & sInitials & "' " _
             & "WHERE (RUNREF='" & sPart & "' AND " _
             & "RUNNO=" & lRunno & ")"
      clsADOCon.ExecuteSQL sSql
      
      'LOTS
      If sLotNumber <> "" Then
         sSql = "UPDATE LohdTable SET " _
                & "LOTDATECOSTED='" & vAdate & "'," _
                & "LOTUNITCOST=" & cRunCost / cYield & "," _
                & "LOTTOTMATL=" & cRunMatl & "," _
                & "LOTTOTLABOR=" & cRunLabor & "," _
                & "LOTTOTEXP=" & cRunExp & "," _
                & "LOTTOTOH=" & cRunOvHd & "," _
                & "LOTTOTHRS=" & cRunHours & " " _
                & "WHERE LOTNUMBER='" & sLotNumber & "'"
         clsADOCon.ExecuteSQL sSql
      End If
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         sSql = "SELECT INTYPE,INNUMBER,INLOTNUMBER FROM InvaTable " _
                & "WHERE (INTYPE=6 AND INLOTNUMBER='" & sLotNumber & "')"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
         If bSqlRows Then
            lInRecord = RdoInv!INNUMBER
            RdoInv.Cancel
            sSql = "UPDATE InvaTable SET " _
                   & "INREF1='CLOSED RUN'," _
                   & "INADATE='" & vAdate & "'," _
                   & "INAMT=" & cRunCost / cYield & "," _
                   & "INTOTMATL=" & cRunMatl & "," _
                   & "INTOTLABOR=" & cRunLabor & "," _
                   & "INTOTEXP=" & cRunExp & "," _
                   & "INTOTOH=" & cRunOvHd & "," _
                   & "INTOTHRS=" & cRunHours & " " _
                   & "WHERE (INTYPE=6 AND INNUMBER=" & lInRecord & ")"
            clsADOCon.ExecuteSQL sSql
         End If
         iLogRecord = iLogRecord + 1
         Print #iLogFile, vbCr
         iLogRecord = iLogRecord + 1
         Print #iLogFile, "Manufacturing Order " & cmbPrt & " Run " & cmbRun & " Was Closed"
         CloseLog
         sMsg = "The Status Was Changed From CO To CL." & vbCrLf _
                & "No Additional Action Can Be Executed."
         MsgBox sMsg, vbInformation, Caption
         bResponse = MsgBox("Would You Like To View The Log Now?", ES_YESQUESTION, Caption)
         If bResponse = vbYes Then OpenWebPage sFilePath & "CloseRuns.txt"
         FillCombo
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         
         iLogRecord = iLogRecord + 1
         Print #iLogFile, vbCr
         iLogRecord = iLogRecord + 1
         Print #iLogFile, "* Manufacturing Order " & cmbPrt & " Run " & cmbRun & " Was Not Closed"
         CloseLog
         MsgBox "Couldn't Change The Run To Closed (CL).", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   CloseLog
   Set RdoInv = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "closemo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   
End Sub




Private Function GetRunPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   cmbRun.Clear
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAPRODCODE," _
          & "PASTDCOST FROM PartTable WHERE PARTREF='" & Compress(cmbPrt) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblLvl = Format(!PALEVEL, "0")
         lblCode = "" & Trim(!PAPRODCODE)
         cStdCost = Format(!PASTDCOST, ES_QuantityDataFormat)
         .Cancel
         GetRunPart = 1
      End With
   Else
      GetRunPart = 0
      lblLvl = ""
      lblCode = ""
      lblDsc = "*** Part Number Wasn't Found ****"
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrunpart"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function CheckInvJournal() As Byte
   Dim b As Byte
   sJournalID = GetOpenJournal("IJ", Format$(txtDte, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   If b = 0 Then
      MsgBox "There Is No Open Inventory Journal For The Period.", _
         vbExclamation, Caption
      CheckInvJournal = 0
      Sleep 500
      Unload Me
   Else
      CheckInvJournal = 1
   End If
   
End Function

'11/11/04 Allocated Purchase Orders

Private Sub GetUnInvoicedPoItems()
   Dim RdoInv As ADODB.Recordset
   Dim sInvoiced As String
   Dim sReceived As String
   Dim sFlag As String
   On Error Resume Next
   iLogRecord = iLogRecord + 1
   Print #iLogFile, vbCr
   iLogRecord = iLogRecord + 1
   If optInv.Value = vbChecked Then
      Print #iLogFile, "Unreceived,Uninvoiced Purchase Orders (Allocations):"
   Else
      Print #iLogFile, "Unreceived Purchase Orders (Allocations):"
   End If
   sSql = "SELECT PINUMBER,PITYPE,PIITEM,PIREV,PIPART,PIRUNPART,PIRUNNO,PIAQTY," _
          & " PARTREF,PARTNUM FROM PoitTable,PartTable " _
          & "WHERE (PIRUNPART='" & Compress(cmbPrt) & "' AND PIRUNNO=" _
          & Val(cmbRun) & " AND PIAQTY=0 AND PITYPE<>16) AND PARTREF=PIPART"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
   If bSqlRows Then
      With RdoInv
         Do Until .EOF
            sFlag = ""
            sReceived = ""
            sInvoiced = ""
            If !PITYPE = 14 Then
               sFlag = "* "
               bCantClose = 1
               sReceived = "Rec:No"
            End If
            If !PITYPE = 15 Then sReceived = "Rec:Yes"
            If optInv.Value = vbChecked Then
               If !PITYPE = 17 Then
                  sInvoiced = "Inv:Yes"
               Else
                  bCantClose = 1
                  sFlag = "* "
                  sInvoiced = "Inv:No"
               End If
            End If
            iLogRecord = iLogRecord + 1
            Print #iLogFile, sFlag & Format(!PINUMBER, "00000"), !PIITEM & !PIREV, _
                                            !PartNum, sReceived, sInvoiced
            .MoveNext
         Loop
         .Cancel
      End With
   Else
      iLogRecord = iLogRecord + 1
      If optInv.Value = vbChecked Then
         Print #iLogFile, "No Unreceived,Uninvoiced Items Found"
      Else
         Print #iLogFile, "No Unreceived Items Found"
      End If
   End If
   Set RdoInv = Nothing
End Sub

Public Sub CloseLog()
   On Error Resume Next
   If iLogFile > 0 Then Close #iLogFile
   iLogFile = 0
   
End Sub

Private Sub GetPickList()
   Dim RdoPck As ADODB.Recordset
   Dim iRow As Integer
   Dim sFlag As String
   Dim sPicked As String
   
   Erase sPartLots
   iTotalPicks = 0
   iLogRecord = iLogRecord + 1
   Print #iLogFile, vbCr
   iLogRecord = iLogRecord + 1
   Print #iLogFile, "Manufacturing Order Picks:"
   '5/24/05
   If optExp.Value = vbUnchecked Then
      sSql = "SELECT PKPARTREF,PKTYPE,PKAQTY,PKMOPART,PKMORUN FROM MopkTable " _
             & "WHERE (PKTYPE<>12 AND PKMOPART='" & Compress(cmbPrt) & "' AND " _
             & "PKMORUN=" & Val(cmbRun) & ")"
   Else
      sSql = "SELECT PARTREF,PALEVEL,PKPARTREF,PKTYPE,PKAQTY,PKMOPART,PKMORUN " _
             & "FROM MopkTable,PartTable WHERE (PARTREF=PKPARTREF AND PKTYPE<>12 " _
             & "AND PKMOPART='" & Compress(cmbPrt) & "' AND PKMORUN=" & Val(cmbRun) _
             & " AND PALEVEL<>5)"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPck, ES_FORWARD)
   If bSqlRows Then
      With RdoPck
         Do Until .EOF
            iRow = iRow + 1
            GetPartNumber !PKPARTREF, iRow
            sPicked = 0
            sFlag = ""
            If !PKTYPE = 9 Or !PKTYPE = 23 Then
               sFlag = "* "
               bCantClose = 1
               sPicked = "Not Picked"
            End If
            If !PKTYPE = 10 Then
               sPicked = Format$(!PKAQTY, "#####0.000") & " Picked"
            End If
            If !PKAQTY = 0 Then
               If optUnpicked.Value = vbUnchecked Then
                  sFlag = "* "
                  bCantClose = 1
                  sPicked = Format$(!PKAQTY, "#####0.000") & " Not Picked (Failed)"
               Else
                  sPicked = Format$(!PKAQTY, "#####0.000") & " Not Picked (Allowed)"
               End If
            End If
            sPartLots(iRow, 3) = Format$(!PKAQTY, "#####0.000")
            iLogRecord = iLogRecord + 1
            Print #iLogFile, sFlag & sPartLots(iRow, 0) & "Qty: " & sPicked
            .MoveNext
         Loop
         iTotalPicks = iRow
         .Cancel
      End With
   Else
      iLogRecord = iLogRecord + 1
      Print #iLogFile, "No Manufacturing Order Picks "
   End If
   Set RdoPck = Nothing
End Sub

'Reserves the material for later use

Public Sub GetPartNumber(PartNumber As String, iRow As Integer)
   Dim RdoGet As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT PARTREF,PARTNUM,PALOTTRACK,PASTDCOST FROM PartTable " _
          & "WHERE PARTREF='" & PartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         sPartLots(iRow, 0) = "" & Trim(!PartNum)
         If bLotsOn Then
            sPartLots(iRow, 1) = Trim(str$(!PALOTTRACK))
         Else
            sPartLots(iRow, 1) = "0"
         End If
         sPartLots(iRow, 2) = Trim(str$(!PASTDCOST))
         .Cancel
      End With
   End If
   Set RdoGet = Nothing
End Sub

'Test Invoicing

Public Sub CheckInvoicing()
   Dim RdoInv As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT COVERIFYINVOICES FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
   If bSqlRows Then optInv.Value = RdoInv!COVERIFYINVOICES
   RdoInv.Cancel
   Set RdoInv = Nothing
   
End Sub

Private Sub SaveSettings()
   SaveSetting "Esi2000", "EsiProd", "diaScncl", Trim(str$(optUnpicked.Value))
   SaveSetting "Esi2000", "EsiProd", "diaScncla", Trim(str$(optExp.Value))
   
End Sub

Private Sub GetSettings()
   optUnpicked.Value = GetSetting("Esi2000", "EsiProd", "diaScncl", Trim(str$(optUnpicked.Value)))
   optExp.Value = GetSetting("Esi2000", "EsiProd", "diaScncla", Trim(str$(optExp.Value)))
   
End Sub
