VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form ShopSHe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Order Completions"
   ClientHeight    =   4710
   ClientLeft      =   1740
   ClientTop       =   1170
   ClientWidth     =   7905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton optDis 
      Height          =   320
      Left            =   6840
      Picture         =   "ShopSHe04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Display The Report"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   490
   End
   Begin VB.CommandButton optPrn 
      Height          =   320
      Left            =   7320
      Picture         =   "ShopSHe04a.frx":017E
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Print The Report"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   490
   End
   Begin VB.ComboBox lblPrinter 
      Height          =   315
      Left            =   1680
      TabIndex        =   39
      Top             =   3840
      Width           =   4455
   End
   Begin VB.CheckBox optLabel 
      Height          =   255
      Left            =   5880
      TabIndex        =   38
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chkPartialCompletion 
      Height          =   255
      Left            =   7200
      TabIndex        =   4
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   2460
      Width           =   495
   End
   Begin VB.CheckBox optExp 
      Height          =   255
      Left            =   3000
      TabIndex        =   34
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3480
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.Frame z2 
      Height          =   30
      Left            =   120
      TabIndex        =   33
      Top             =   1920
      Width           =   7692
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHe04a.frx":0308
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optLot 
      Alignment       =   1  'Right Justify
      Caption         =   "Lot Tracked     "
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Tag             =   "4"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCom 
      Caption         =   "C&omplete"
      Height          =   315
      Left            =   6900
      TabIndex        =   5
      ToolTipText     =   "Complete The Selected Manufacturing Order"
      Top             =   2760
      Width           =   875
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Actual Quantity"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   720
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   6840
      TabIndex        =   1
      ToolTipText     =   "Select Run Number"
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6960
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8160
      Top             =   3360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4710
      FormDesignWidth =   7905
   End
   Begin VB.Label lblPrePickMO 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   240
      TabIndex        =   43
      Top             =   4320
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "Label Printer"
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Print Label"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   13
      Left            =   4440
      TabIndex        =   37
      ToolTipText     =   "Print MO Label"
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Partial Completion?"
      ForeColor       =   &H00000000&
      Height          =   465
      Index           =   12
      Left            =   6840
      TabIndex        =   36
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore Part Type 5's (Expendables)"
      ForeColor       =   &H00000000&
      Height          =   228
      Index           =   11
      Left            =   240
      TabIndex        =   35
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3480
      Width           =   3492
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   6
      Left            =   5520
      TabIndex        =   31
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblRunQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6840
      TabIndex        =   30
      ToolTipText     =   "Original Run Quantity"
      Top             =   1440
      Width           =   875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Quantity"
      Height          =   255
      Index           =   10
      Left            =   5520
      TabIndex        =   29
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblScrap 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5880
      TabIndex        =   28
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblRework 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5880
      TabIndex        =   27
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scrap"
      Height          =   255
      Index           =   9
      Left            =   4440
      TabIndex        =   26
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rework"
      Height          =   255
      Index           =   8
      Left            =   4440
      TabIndex        =   25
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Partial Complete"
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   24
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblPartial 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5880
      TabIndex        =   23
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblLvl 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2400
      TabIndex        =   21
      ToolTipText     =   "Part Type (Level)"
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   19
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Completed"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   2055
      Width           =   1575
   End
   Begin VB.Label lblTxt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   17
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Complete"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   2415
      Width           =   1575
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      ToolTipText     =   "Remaining Quantity"
      Top             =   2760
      Width           =   1075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mo quantity left"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2880
      TabIndex        =   13
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblSch 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5880
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sched Complete"
      Height          =   255
      Index           =   14
      Left            =   4440
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   760
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   9
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6840
      TabIndex        =   8
      ToolTipText     =   "Run Status"
      Top             =   1080
      Width           =   876
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "ShopSHe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES2000 is the property of                                ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***

Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter


Dim bGoodPart As Byte
Dim bGoodMo As Byte
Dim bOnLoad As Byte
Dim bPartLot As Byte
Dim bPartialCompletion As Boolean
Dim bDisLostFocus As Boolean

Dim lRunno As Long

Dim cRunExp As Currency
Dim cRunHours As Currency
Dim cRunLabor As Currency
Dim cRunMatl As Currency
Dim cRunOvHd As Currency
Dim cStdCost As Currency
Dim cRunCost As Currency

Dim sPartNumber As String
Dim sCreditAcct As String
Dim sDebitAcct As String
'WIP
'Dim sInvLabAcct As String
'Dim sInvMatAcct As String
'Dim sInvExpAcct As String
'Dim sInvOhdAcct As String

Dim sCgsLabAcct As String
Dim sCgsMatAcct As String
Dim sCgsExpAcct As String
Dim sCgsOhdAcct As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Public quantityPerLabel As Currency
Public totalQuantity As Currency

'Complete


'Dim cCompletionQty As Currency
'Dim cSuggestedQty As Currency

Private Sub GetRunCosts()
   Dim RdoCst As ADODB.Recordset
   cRunExp = 0
   cRunHours = 0
   cRunLabor = 0
   cRunMatl = 0
   cRunOvHd = 0
   
   sProcName = "getruncosts"
   sSql = "SELECT RUNREF,RUNNO,RUNCOST,RUNOHCOST,RUNCMATL," _
          & "RUNCEXP,RUNCHRS,RUNCLAB FROM RunsTable WHERE " _
          & "RUNREF='" & sPartNumber & "' AND RUNNO=" & lRunno & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   If bSqlRows Then
      With RdoCst
         cRunExp = Format(!RUNCEXP, ES_QuantityDataFormat)
         cRunHours = Format(!RUNCHRS, ES_QuantityDataFormat)
         cRunLabor = Format(!RUNCLAB, ES_QuantityDataFormat)
         cRunMatl = Format(!RUNCMATL, ES_QuantityDataFormat)
         cRunOvHd = Format(!RUNOHCOST, ES_QuantityDataFormat)
         ClearResultSet RdoCst
      End With
   End If
   Set RdoCst = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtQty = "0.000"
   
End Sub

Private Sub cmbPrt_Click()
   bGoodPart = GetPart(False)
   
End Sub

Private Sub cmbPrt_GotFocus()
   cmdCom.Enabled = False
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   bGoodPart = GetPart(False)
   If Not bGoodPart Then
      If (bDisLostFocus = False) And cmbPrt <> "" Then
        MsgBox "Invalid Part Number Selected"
        cmbPrt.SetFocus
      End If
   End If
End Sub


Private Sub cmbRun_Click()
   If Val(cmbRun) > 0 Then bGoodMo = GetRun()
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   If Val(cmbRun) > 0 Then bGoodMo = GetRun()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCom_Click()
   Dim sMsg As String
   Dim bResponse As Byte
   
    If lblStat <> "PC" And Not AllowMOCompletionWhenNotPC Then
      If (lblStat = "RL") Then
        
         bResponse = MsgBox("You Cannot Complete an MO that is Release (RL) Status." & vbCrLf & "Do You Wish To Complete The MO Anyway?", ES_NOQUESTION, Caption)
         If bResponse = vbNo Then
            CancelTrans
            Exit Sub
         End If
      Else
        MsgBox "You Cannot Complete an MO that is not at Pick Complete (PC) Status" & vbCrLf & "(See System Settings for More Information)", vbInformation, Caption
        CancelTrans
        Exit Sub
      End If
    End If
   
   If lblStat = "PL" Or lblStat = "PP" Then
      bResponse = MsgBox("The Pick Is Not Complete At " & lblStat & "." _
                  & vbCr _
                  & "Do You Wish To Complete The MO Anyway?", _
                  ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         CancelTrans
         Exit Sub
      End If
   End If
   
   ' don't allow completion if open time charges
   Dim rdoCharges As ADODB.Recordset
   sSql = "select distinct rtrim(PREMFSTNAME) + ' ' + rtrim(PREMLSTNAME) as Name," & vbCrLf _
    & "'Op ' + rtrim((" & vbCrLf _
    & " select cast(ISOP as VARCHAR(4)) + ' '  from IstcTable chg2" & vbCrLf _
    & " where chg2.ISMO = chg1.ISMO and chg2.ISRUN = chg1.ISRUN" & vbCrLf _
    & " order by ISOP" & vbCrLf _
    & " for XML PATH('')" & vbCrLf _
    & ")) as Ops" & vbCrLf _
    & "from IstcTable chg1" & vbCrLf _
    & "join EmplTable emp on emp.PREMNUMBER = chg1.ISEMPLOYEE" & vbCrLf _
    & "where ISMO = '" & Compress(cmbPrt) & "' and ISRUN = " & cmbRun.Text
    '& "where ISMO = '111T50115171' and ISRUN = 10"

   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCharges)
   If bSqlRows Then
        Dim msg As String
        msg = "Cannot complete MO until time charges are closed:" & vbCrLf
        With rdoCharges
            msg = msg & !Name & " " & !Ops & vbCrLf
            .MoveNext
        End With
        MsgBox msg
        Exit Sub
   End If
   Dim cCompletionQty As Currency
   Dim cSuggestedQty As Currency
   bPartialCompletion = CBool(chkPartialCompletion)
   cCompletionQty = Val(txtQty)
   cSuggestedQty = Val(lblQty)
   
   If cCompletionQty < 0 Then
      MsgBox "You cannot complete for a negative quantity."
      'CancelTrans
      Exit Sub
   ElseIf bPartialCompletion And cCompletionQty = 0 Then
      MsgBox "Partial completion and no quantity specified.  Nothing to do."
      'CancelTrans
      Exit Sub
   ElseIf bPartialCompletion And cCompletionQty = cSuggestedQty Then
      MsgBox ("Can not have partial completion with a quantity (" & cCompletionQty & ") equal " & _
         " to the remaining suggested quantity (" & cSuggestedQty & ")." & vbCr & _
         " Please uncheck the Partial Completion Flag.")
      ' Allow users to conmplete all the qty and still keep runstatus as PC
      Exit Sub
   
   ElseIf bPartialCompletion And cCompletionQty > cSuggestedQty Then
      
      If (AllowOverMOQty = True) Then
         bResponse = MsgBox("You have requested to partial completion quantity (" & cCompletionQty & ") greater than" & _
            " the remaining suggested quantity (" & cSuggestedQty & ")." & vbCr & _
            " Is this Correct ?", _
         ES_YESQUESTION, Caption)
      Else
         MsgBox ("Can not have partial completion quantity (" & cCompletionQty & ") greater than" & _
            " or equal to the remaining suggested quantity (" & cSuggestedQty & ")." & vbCr & _
            " Please revise the MO quantity.")
         Exit Sub
      End If
   ElseIf bPartialCompletion Then
      bResponse = MsgBox("You have requested a partial completion with a quantity of " & cCompletionQty _
         & ". Is this correct ?", _
         ES_YESQUESTION, Caption)
   
   ElseIf cCompletionQty > cSuggestedQty Then
      
      If (AllowOverMOQty = True) Then
         bResponse = MsgBox("You have requested to MO completion quantity (" & cCompletionQty & ") greater than" & _
            " the remaining suggested quantity (" & cSuggestedQty & ")." & vbCr & _
            " Is this Correct ?", _
         ES_YESQUESTION, Caption)
      Else
         MsgBox ("Can not have MO completion quantity (" & cCompletionQty & ") greater than" & _
            " the remaining suggested quantity (" & cSuggestedQty & ")." & vbCr & _
            " Please revise the MO quantity.")
         Exit Sub
      End If
      
   
   ElseIf cCompletionQty = 0 Then
      bResponse = MsgBox("You have requested a full complete with no quantity specified.  Is this correct?", _
         ES_YESQUESTION, Caption)
   ElseIf cCompletionQty <> cSuggestedQty Then
      bResponse = MsgBox("You have requested a completion with a quantity (" & cCompletionQty & ") not equal to" & _
         " the remaining suggested quantity (" & cSuggestedQty & ").  Is this correct?", _
         ES_YESQUESTION, Caption)
   Else
      bResponse = MsgBox("You have requested a full completion with a quantity of " & cCompletionQty _
         & ".  Is this correct?", _
         ES_YESQUESTION, Caption)
   End If
   
   If bResponse <> vbYes Then
      'CancelTrans
      On Error Resume Next
      txtQty.SetFocus
      Exit Sub
   End If
   'GetPreviousCompletions
   On Error GoTo DiaErr1
   sPartNumber = Compress(cmbPrt)
   lRunno = Val(cmbRun)
   MouseCursor ccHourglass
'   Dim mo As New ClassMO
'   mo.PartNumber = sPartNumber
'   mo.RunNumber = lRunno
'
'   mo.GetRunCosts
'   cRunHours = mo.RunHours
'   cRunOvHd = mo.RUNOVHD
'   cRunLabor = mo.RunLabor
'   cRunMatl = mo.RUNMATL
'   cRunExp = mo.RUNEXP
'   cRunCost = mo.RUNCOST
'
'   MouseCursor 13
'   GetWipAccounts
   CompleteMo
   chkPartialCompletion.Value = vbUnchecked
   
   If (optLabel.Value = vbChecked) Then
      
      Dim RdoLot As ADODB.Recordset
      Dim strMOPart As String
      Dim strRun As String
      Dim strLotNum As String
      Dim strUserLot As String
      Dim strOrgQty As String
      Dim strComQty As String
      Dim strLotLoc As String
      
      strMOPart = sPartNumber
      strRun = lRunno

      strComQty = ""
      If cCompletionQty = 0 Then
         sSql = "SELECT SUM(LOTORIGINALQTY) LOTORIGINALQTY" _
               & " FROM lohdtable, LoitTable" _
            & " WHERE lohdtable.lotNumber = LoitTable.LOINUMBER" _
               & " AND LOITYPE = 6 AND lohdtable.LOTMOPARTREF = '" & strMOPart & "'" _
               & " AND LOTMORUNNO = " & strRun

         bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
         If bSqlRows Then
            With RdoLot
               strComQty = !LOTORIGINALQTY
               ClearResultSet RdoLot
            End With
         End If
         Set RdoLot = Nothing
      End If
      
      
      sSql = "SELECT TOP(1) LOTNUMBER, LOTUSERLOTID, LOTORIGINALQTY, LOTLOCATION " _
            & " FROM lohdtable, LoitTable" _
         & " WHERE lohdtable.lotNumber = LoitTable.LOINUMBER" _
            & " AND LOITYPE = 6 AND lohdtable.LOTMOPARTREF = '" & strMOPart & "'" _
            & " AND LOTMORUNNO = " & strRun & " ORDER BY LOTADATE DESC"
   
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
      
      If bSqlRows Then
         With RdoLot
            strLotNum = !lotNumber
            strUserLot = !LOTUSERLOTID
            If cCompletionQty = 0 Then
               strOrgQty = strComQty
            Else
               strOrgQty = !LOTORIGINALQTY
            End If
            
            strLotLoc = !LOTLOCATION
            ClearResultSet RdoLot
         End With
      End If
      Set RdoLot = Nothing
      
      
      Load ShopSHe04b
      ShopSHe04b.lblPartNo = strMOPart
      ShopSHe04b.lblRunNo = strRun
      ShopSHe04b.lblLotNum = strLotNum
      ShopSHe04b.lblUserLotNum = strUserLot
      ShopSHe04b.lblLocation = strLotLoc
      ShopSHe04b.lblQty = strOrgQty
      ShopSHe04b.txtQtyPerLabel = strOrgQty
      totalQuantity = strOrgQty
      Set ShopSHe04b.ParentForm = Me
      bDisLostFocus = True
      ShopSHe04b.Show vbModal
      If Me.quantityPerLabel > 0 Then
         'MsgBox "print " & Me.quantityPerLabel & " per label"
         PrintLabels strMOPart, strRun, Me.totalQuantity, Me.quantityPerLabel
      End If
      cmbRun = lRunno
      bDisLostFocus = False
                  
   End If
   
   MouseCursor ccArrow
   ' MM FillRuns Me, "NOT LIKE 'C%'"
   bGoodPart = GetPart(False)
   ' Reset the vlaue
   ' If the Run is the last one,....
   'then we need to move to the next partnumber.
   Dim Index As Integer
   If (bGoodPart = False) Then
      If (cmbPrt.ListIndex = -1) Then
         Index = 0
      Else
         Index = cmbPrt.ListIndex
      End If
      cmbPrt.RemoveItem (Index)
      
      cmbPrt.ListIndex = cmbPrt.ListIndex + 1
      bGoodPart = GetPart(False)
      ' MM FillRuns Me, "NOT LIKE 'C%'"
   Else
      cmbRun = lRunno
   End If
   
   
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintLabels(strMOPart As String, strRun As String, totalQuantity As Currency, quantityPerLabel As Currency)
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   Dim quantityLeft As Currency
   quantityLeft = totalQuantity
   optPrn.Value = True
   Do
      
      sCustomReport = GetCustomReport("prdshe01")
      Set cCRViewer = New EsCrystalRptViewer
      cCRViewer.Init
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
      aFormulaName.Add "Quantity"
       
      If quantityLeft > quantityPerLabel Then
         aFormulaValue.Add CStr("'" & quantityPerLabel & "'")
      Else
         aFormulaValue.Add CStr("'" & quantityLeft & "'")
      End If
       
      cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

      sSql = "{lohdtable.LOTMORUNNO} = " & strRun & " AND {lohdtable.LOTMOPARTREF} = '" & strMOPart & "'" _
            & " AND {loitTable.LOITYPE} = 6"
            
      cCRViewer.SetReportSelectionFormula (sSql)
      cCRViewer.CRViewerSize Me
      cCRViewer.ShowGroupTree False
      cCRViewer.SetDbTableConnection
   
      cCRViewer.OpenCrystalReportObject Me, aFormulaName, 1, True
      
      cCRViewer.ClearFieldCollection aFormulaName
      cCRViewer.ClearFieldCollection aFormulaValue
      Set cCRViewer = Nothing
      
      quantityLeft = quantityLeft - quantityPerLabel
   
   Loop While quantityLeft > 0
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "PrintLabels"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4104
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MouseCursor 0
   Dim b As Byte
   Dim X As Printer
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillRuns Me, "NOT LIKE 'C%'"
      
    '4/4/17 asked by IMAINC not to initially populate # 53
    cmbPrt.ListIndex = -1
    cmbPrt.Text = ""
    cmbRun.ListIndex = -1
    cmbRun.Text = ""
      
      bGoodPart = GetPart(True)
      On Error Resume Next
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
      
      
      'populate printer selection combo
      For Each X In Printers
         If Left(X.DeviceName, 9) <> "Rendering" Then
            lblPrinter.AddItem X.DeviceName
         End If
      Next
      
      On Error Resume Next
      
      Dim sDefaultPrinter As String
      If lblPrinter.ListCount > 0 Then
         sDefaultPrinter = lblPrinter.List(0)
      End If
      
      lblPrinter.Text = GetSetting("Esi2000", "EsiProd", "MOLabelPrinter", sDefaultPrinter)
      lblPrePickMO = ""
      lblPrePickMO.Visible = False
      
      bOnLoad = 0
   End If
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS," _
          & "PALEVEL,PARUN,PASTDCOST,PAPRODCODE,PALOTTRACK," _
          & "RUNREF,RUNNO,RUNSTATUS,RUNQTY,RUNSCHED,RUNPARTIALQTY " _
          & "FROM PartTable,RunsTable WHERE PARTREF= ? " _
          & "AND PARTREF=RUNREF AND RUNSTATUS NOT LIKE 'C%'"
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   
   AdoQry.Parameters.Append AdoParameter
   
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   bDisLostFocus = False
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   FormUnload
   
   SaveSetting "Esi2000", "EsiProd", "MOLabelPrinter", lblPrinter.Text
   
   Set ShopSHe04a = Nothing
   
End Sub




Private Function GetRun() As Byte
   Dim RdoRun As ADODB.Recordset
   sPartNumber = Compress(cmbPrt)
   
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNQTY,RUNSTATUS,RUNSCHED,RUNPARTIALQTY," _
          & "RUNREWORK,RUNSCRAP,RUNREMAININGQTY,ISNULL(RUNPREPKMOREF, '') RUNPREPKMOREF," _
          & "ISNULL(RUNPREPKNO, '') RUNPREPKNO FROM RunsTable " _
          & "WHERE RUNREF='" & sPartNumber & "' AND RUNNO=" & cmbRun _
          & " AND RUNSTATUS NOT LIKE 'C%'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun)
   If bSqlRows Then
      With RdoRun
         GetRun = True
         lblRunQty = Format(!RUNQTY, ES_QuantityDataFormat)
         lblQty = Format(!RUNREMAININGQTY, ES_QuantityDataFormat)
         If !RUNREMAININGQTY < 0 Then
            txtQty = Format(0, ES_QuantityDataFormat)
         Else
            txtQty = Format(!RUNREMAININGQTY, ES_QuantityDataFormat)
         End If
         lblPartial = Format(!RUNPARTIALQTY, ES_QuantityDataFormat)
         lblRework = Format(!RUNREWORK, ES_QuantityDataFormat)
         lblScrap = Format(!RUNSCRAP, ES_QuantityDataFormat)
         lblStat = "" & !RUNSTATUS
         lblSch = "" & Format(!runSched, "mm/dd/yyyy")
         
         If (Trim(!RUNPREPKMOREF) <> "") Then
            lblPrePickMO.Visible = True
            lblPrePickMO = "Prepick To - " & Trim(!RUNPREPKMOREF) & " And Run No - " & Trim(!RUNPREPKNO)
         Else
            lblPrePickMO = ""
            lblPrePickMO.Visible = False
         End If
         
         Select Case lblStat
            Case "SC"
               lblTxt = "Scheduled"
            Case "PL", "PP"
               lblTxt = "Pick Is Not Complete"
            Case "PC"
               lblTxt = "Pick Is Complete"
         End Select
         txtQty.Enabled = True
         ClearResultSet RdoRun
      End With
   Else
      lblStat = ""
      lblQty = "0.000"
      txtQty = "0.000"
      lblTxt = ""
      sPartNumber = ""
      txtQty.Enabled = False
      MouseCursor 0
      'MsgBox "Run Wasn't Found. May Be CO,CL or CA..", vbInformation, Caption
      GetRun = False
   End If
   On Error Resume Next
   Set RdoRun = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetPart(bOnLoad As Byte) As Byte
   Dim RdoPrt As ADODB.Recordset
   If bOnLoad Then
        GetPart = False
        Exit Function
    End If
   
   sPartNumber = Compress(cmbPrt)
   
   On Error GoTo DiaErr1
   cmbRun.Clear
   cmdCom.Enabled = False
   AdoQry.Parameters(0).Value = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoPrt, AdoQry, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         cmbRun = Format(!Runno, "####0")
         cStdCost = Format(!PASTDCOST, ES_QuantityDataFormat)
         lblQty = Format(!RUNQTY, ES_QuantityDataFormat)
         lblStat = "" & !RUNSTATUS
         lblDsc = "" & Trim(!PADESC)
         lblUom = "" & Trim(!PAUNITS)
         lblSch = "" & Format(!runSched, "mm/dd/yyyy")
         lblLvl = Format$(!PALEVEL, "0")
         lblCode = "" & Trim(!PAPRODCODE)
         bPartLot = !PALOTTRACK
         optLot.Value = !PALOTTRACK
         GetPart = True
         Do Until .EOF
            AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoPrt
      End With
      On Error Resume Next
'      '4/4/17 asked by IMAINC not to initially populate # 53
'      If bOnLoad Then
''        cmbPrt.ListIndex = -1
''        cmbPrt.Text = ""
''        cmbRun.ListIndex = -1
''        cmbRun.Text = ""
'      End If
      If cmbRun.ListCount > 0 Then cmbRun.ListIndex = 0
      If GetPreferenceValue("AutoSelectLastRun") = "1" Then cmbRun = cmbRun.List(cmbRun.ListCount - 1)
      lblPrePickMO = ""
      lblPrePickMO.Visible = False

      bGoodMo = GetRun()
   Else
      optLot.Value = vbUnchecked
      sPartNumber = ""
      bPartLot = 0
      cStdCost = 0
      'cmbRun = "0"
      lblUom = ""
      lblDsc = ""
      lblSch = ""
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function





Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDateEx(txtDte)
   
End Sub


Private Sub txtQty_Click()
   SelectFormat Me
   If bGoodMo Then cmdCom.Enabled = True Else cmdCom.Enabled = False
   
End Sub

Private Sub txtQty_GotFocus()
   txtQty_Click
   
End Sub


Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   'txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   txtQty = Format(txtQty, ES_QuantityDataFormat)
End Sub
'
'Private Sub CompleteMo()
'   Dim bByte As Byte
'   Dim bResponse As Byte
'   Dim bLots As Byte
'
'   Dim iLength As Integer
'   Dim iOps As Integer
'
'   Dim lCOUNTER As Long
'   Dim lSysCount As Long
'
'   Dim cComQty As Currency
'   Dim cQuantity As Currency
'   'Dim cRunCost As Currency
'
'   Dim sLotNumber As String
'   Dim sMsg As String
'   Dim sPartNum As String
'   Dim sStatus As String
'
'   Dim vDate As Variant
'
'   cComQty = Format(Val(txtQty), ES_QuantityDataFormat)
'   cQuantity = Format(Val(lblQty), ES_QuantityDataFormat)
'   vDate = Format(ES_SYSDATE, "mm/dd/yy hh:mm")
'
'   On Error GoTo DiaErr1
'
'   'if partial completion, leave status as it is
'   If bPartialCompletion Then
'      sStatus = Me.lblStat
'   Else
'      sStatus = "CO"
'   End If
'
'   iOps = GetOpCompletions()
'   MouseCursor 0
'   If iOps > 0 Then
'      bResponse = MsgBox("This Manufacturing Order Contains" & vbCr _
'                  & "(" & iOps & ") Operations That Are Not Complete. " & vbCr _
'                  & "Continue Anyway?", _
'                  ES_NOQUESTION, Caption)
'      If bResponse = vbNo Then
'         CancelTrans
'         Exit Sub
'      End If
'   End If
'   MouseCursor 13
'   cmdCom.Enabled = False
'   iLength = Len(Trim(Str(cmbRun)))
'   iLength = 5 - iLength
'   If iLength < 0 Then iLength = 0
'
'   On Error GoTo DiaErr1
'
'   bResponse = GetPartAccounts(sPartNumber, sCreditAcct, sDebitAcct)
'   cRunCost = Format((cRunOvHd + cRunMatl + cRunLabor + cRunExp), "######0.000")
'   Err = 0
'
'   'Check lots
'   bLots = CheckLotStatus()
'   lCOUNTER = GetLastActivity() + 1
'   sLotNumber = GetNextLotNumber()
'   lSysCount = lCOUNTER
'
'   clsAdoCon.begintrans
'   lblPartial mo.GetPreviousCompletions
'   sPartNum = Trim(cmbPrt)
'   If Len(sPartNum) > 25 Then Compress (sPartNum)
'   If cComQty > 0 Then
'      sSql = "UPDATE PartTable SET PAQOH=PAQOH+" & cComQty & "," _
'             & "PALOTQTYREMAINING=PALOTQTYREMAINING+" & cComQty & " " _
'             & "WHERE PARTREF='" & sPartNumber & "'"
'      clsAdoCon.ExecuteSQL sSql
'   End If
'
'   'create ia & lot records if the quantity is > 0
'   'or if qty = 0, this is a final completion, and there are no other completions
'   'and there are no other
'   Dim rdo As ADODB.recordset
'   Dim totalMoQty As Currency
'   totalMoQty = 0
'   If cComQty = 0 Then
'      sSql = "select RUNYIELD from RunsTable" & vbCrLf _
'         & "WHERE RUNREF='" & sPartNumber & "' AND " _
'         & "RUNNO=" & Val(cmbRun) & " "
'      If GetDataSet(rdo) Then
'         totalMoQty = CCur(rdo.fields(0).Value)
'      End If
'   End If
'
'   If cComQty > 0 Or (cComQty = 0 And totalMoQty = 0) Then
'      sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
'             & "INPDATE,INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," _
'             & "INMOPART,INMORUN,INTOTMATL,INTOTLABOR,INTOTEXP," _
'             & "INTOTOH,INTOTHRS,INWIPLABACCT,INWIPMATACCT," _
'             & "INWIPOHDACCT,INWIPEXPACCT,INNUMBER,INLOTNUMBER,INUSER) " _
'             & "VALUES(6,'" & sPartNumber & "','COMPLETED RUN'," _
'             & "'RUN " & String(iLength, Chr(32)) & Trim(Str(cmbRun)) & "'," _
'             & "'" & txtDte & "','" & vDate & "'," & cComQty & "," & cComQty & "," _
'             & cRunCost & ",'" & sCreditAcct & "','" & sDebitAcct & "','" _
'             & sPartNumber & "'," & Val(cmbRun) & ","
'      sSql = sSql & cRunMatl & "," & cRunLabor & "," _
'             & cRunExp & "," & cRunOvHd & "," & cRunHours & ",'" _
'             & sInvLabAcct & "','" & sInvMatAcct & "','" _
'             & sInvExpAcct & "','" & sInvOhdAcct & "'," & lCOUNTER & ",'" _
'             & sLotNumber & "','" & sInitials & "')"
'      clsAdoCon.ExecuteSQL sSql
'
'      If bPartialCompletion Then
'         sMsg = "MO PA-"
'      Else
'         sMsg = "MO CO-"
'      End If
'      sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
'             & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
'             & "LOTUNITCOST,LOTMOPARTREF,LOTMORUNNO," _
'             & "LOTTOTMATL,LOTTOTLABOR,LOTTOTEXP,LOTTOTOH,LOTTOTHRS) " _
'             & "VALUES('" _
'             & sLotNumber & "','" & sMsg & sPartNum & " R" & Trim(Val(cmbRun)) _
'             & "','" & sPartNumber & "','" & vDate & "'," & cComQty & "," & Trim(Str(cComQty)) _
'             & "," & cRunCost & ",'" & sPartNumber & "'," _
'             & Val(cmbRun) & "," & cRunMatl & "," & cRunLabor & "," _
'             & cRunExp & "," & cRunOvHd & "," & cRunHours & ")"
'      clsAdoCon.ExecuteSQL sSql
'      If bPartialCompletion Then
'         sMsg = "MO Run Partial Comp"
'      Else
'         sMsg = "MO Run Completion"
'      End If
'      sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
'             & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
'             & "LOIMOPARTREF,LOIMORUNNO,LOIACTIVITY,LOICOMMENT) " _
'             & "VALUES('" _
'             & sLotNumber & "',1,6,'" & sPartNumber _
'             & "','" & txtDte & "'," & Trim(Str(cComQty)) _
'             & ",'" & sPartNumber & "'," & Val(cmbRun) & "," _
'             & lCOUNTER & ",'MO Run Completion')"
'      clsAdoCon.ExecuteSQL sSql
'   End If
'
'   'partial completion
'   If bPartialCompletion Then
'      Err.Clear
'      sSql = "UPDATE RunsTable SET " _
'             & "RUNYIELD=" & cComQty & "," _
'             & "RUNPARTIALQTY=RUNPARTIALQTY+" & cComQty & "," _
'             & "RUNPARTIALDATE='" & txtDte & "'," _
'             & "RUNREMAININGQTY=RUNREMAININGQTY-" & cComQty & "," _
'             & "RUNLOTNUMBER='" & sLotNumber & "' " _
'             & "WHERE RUNREF='" & sPartNumber & "' AND " _
'             & "RUNNO=" & Val(cmbRun) & " "
'      clsAdoCon.ExecuteSQL sSql
'
'   'full completion
'   Else
'      sSql = "UPDATE RnopTable SET OPCOMPDATE='" & txtDte & "'," _
'             & "OPCOMPLETE=1 WHERE OPREF='" & sPartNumber & "' AND " _
'             & "OPRUN=" & Val(cmbRun) & " AND OPCOMPLETE=0 "
'      clsAdoCon.ExecuteSQL sSql
'      sSql = "UPDATE RunsTable SET RUNCOMPLETE='" & txtDte & "'," _
'             & "RUNYIELD=" & cComQty & ",RUNSTATUS='" & sStatus & "'," _
'             & "RUNCOST=" & cRunCost & "," _
'             & "RUNOHCOST=" & cRunOvHd & "," _
'             & "RUNCMATL=" & cRunMatl & "," _
'             & "RUNCEXP=" & cRunExp & "," _
'             & "RUNCHRS=" & cRunHours & "," _
'             & "RUNCLAB=" & cRunLabor & "," _
'             & "RUNREMAININGQTY=0," _
'             & "RUNLOTNUMBER='" & sLotNumber & "' " _
'             & "WHERE RUNREF='" & sPartNumber & "' AND " _
'             & "RUNNO=" & Val(cmbRun) & " "
'      clsAdoCon.ExecuteSQL sSql
'   End If
'   MouseCursor 0
'
'   Dim mo As New ClassMO
'   Dim sReturnMsg As String
'   sReturnMsg = mo.UpdateMOCosts(sPartNumber, Val(cmbRun), "COMPLETED RUN", _
'             cRunMatl, cRunLabor, cRunExp, cRunOvHd, cRunHours)
'
'   clsAdoCon.CommitTrans
'   AverageCost (sPartNumber)
''   If bLots And bPartLot Then
''      If cComQty > 0 Then
''         MsgBox "The run was completed and lot number " & vbCrLf _
''            & "System: " & sLotNumber & " was created." & vbCrLf _
''            & "You may now edit the lot.", _
''            vbInformation, Caption
''
''         LotEdit.optMo.Value = vbChecked
''         LotEdit.DebitAccount = sDebitAcct
''         LotEdit.CreditAccount = sDebitAcct
''         LotEdit.MONUMBER = Compress(cmbPrt)
''         LotEdit.MORUN = cmbRun
''         LotEdit.INVACTIVITY = lCOUNTER
''         LotEdit.lblOrigQty = cComQty
''
''         LotEdit.RUNCOST = cRunCost
''         LotEdit.RUNOVHD = cRunOvHd
''         LotEdit.RUNMATL = cRunMatl
''         LotEdit.RUNEXP = cRunExp
''         LotEdit.RUNHRS = cRunHours
''         LotEdit.RUNLBR = cRunLabor
''
''         LotEdit.txtLong = "MO Completion"
''         LotEdit.txtlot = "MO CO-" & cmbPrt & " Run " & Trim(Val(cmbRun))
''
''         LotEdit.INVLABACCT = sInvLabAcct
''         LotEdit.INVMATACCT = sInvMatAcct
''         LotEdit.INVEXPACCT = sInvExpAcct
''         LotEdit.INVOHDACCT = sInvOhdAcct
''
''         LotEdit.lblPart = cmbPrt
''         LotEdit.lblDate = Format(ES_SYSDATE, "mm/dd/yy")
''         LotEdit.lblTime = Format(ES_SYSDATE, "hh:mm")
''         LotEdit.lblNumber = sLotNumber
''         LotEdit.Show 1
''      End If
''
''      UpdateWipColumns lSysCount
''      SysMsg "Manufacturing Order Is Complete.", True, Me
''
''      On Error Resume Next
''      cmbRun.Clear
''      MouseCursor 0
''      cmbPrt.SetFocus
''   Else
''      clsAdoCon.RollbackTrans
''      MsgBox "Could Not Complete The Manufacturing Order.", _
''         vbExclamation, Caption
''   End If
'
'   If bLots And bPartLot And cComQty > 0 Then
'      MsgBox "The run was completed and lot number " & vbCrLf _
'         & "System: " & sLotNumber & " was created." & vbCrLf _
'         & "You may now edit the lot.", _
'         vbInformation, Caption
'
'      LotEdit.optMo.Value = vbChecked
'      LotEdit.DebitAccount = sDebitAcct
'      LotEdit.CreditAccount = sDebitAcct
'      LotEdit.MONUMBER = Compress(cmbPrt)
'      LotEdit.MORUN = cmbRun
'      LotEdit.INVACTIVITY = lCOUNTER
'      LotEdit.lblOrigQty = cComQty
'
'      LotEdit.RUNCOST = cRunCost
'      LotEdit.RUNOVHD = cRunOvHd
'      LotEdit.RUNMATL = cRunMatl
'      LotEdit.RUNEXP = cRunExp
'      LotEdit.RUNHRS = cRunHours
'      LotEdit.RUNLBR = cRunLabor
'
'      LotEdit.txtLong = "MO Completion"
'      LotEdit.txtlot = "MO CO-" & cmbPrt & " Run " & Trim(Val(cmbRun))
'
'      LotEdit.INVLABACCT = sInvLabAcct
'      LotEdit.INVMATACCT = sInvMatAcct
'      LotEdit.INVEXPACCT = sInvExpAcct
'      LotEdit.INVOHDACCT = sInvOhdAcct
'
'      LotEdit.lblPart = cmbPrt
'      LotEdit.lblDate = Format(ES_SYSDATE, "mm/dd/yy")
'      LotEdit.lblTime = Format(ES_SYSDATE, "hh:mm")
'      LotEdit.lblNumber = sLotNumber
'      LotEdit.Show 1
'   End If
'
'   UpdateWipColumns lSysCount
'   SysMsg "Manufacturing Order Is Complete.", True, Me
'
'   On Error Resume Next
'   cmbRun.Clear
'   MouseCursor 0
'   cmbPrt.SetFocus
'
''   Else
''      clsAdoCon.RollbackTrans
''      MsgBox "Could Not Complete The Manufacturing Order.", _
''         vbExclamation, Caption
''   End If
'
'   Exit Sub
'
'DiaErr1:
'   sProcName = "completemo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   On Error Resume Next
'   clsAdoCon.RollbackTrans
'   DoModuleErrors Me
'
'End Sub


Private Sub CompleteMo()
   Dim bByte As Byte
   Dim bResponse As Byte
   'Dim bLots As Byte

   Dim iLength As Integer
   Dim iOps As Integer

   Dim lCOUNTER As Long
   Dim lSysCount As Long

   Dim cComQty As Currency
   Dim cQuantity As Currency
   'Dim cRunCost As Currency

   'Dim sLotNumber As String
   Dim sMsg As String
   Dim sPartNum As String
   'Dim sStatus As String

   Dim vDate As Variant

   cComQty = Format(Val(txtQty), ES_QuantityDataFormat)
   cQuantity = Format(Val(lblQty), ES_QuantityDataFormat)
   vDate = Format(ES_SYSDATE, "mm/dd/yy hh:mm")

   On Error GoTo DiaErr1

'   'if partial completion, leave status as it is
'   If bPartialCompletion Then
'      sStatus = Me.lblStat
'   Else
'      sStatus = "CO"
'   End If

   iOps = GetOpCompletions()
   MouseCursor 0
   If iOps > 0 Then
      bResponse = MsgBox("This Manufacturing Order Contains" & vbCr _
                  & "(" & iOps & ") Operations That Are Not Complete. " & vbCr _
                  & "Continue Anyway?", _
                  ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         CancelTrans
         Exit Sub
      End If
   End If
   MouseCursor 13
   cmdCom.Enabled = False
'   iLength = Len(Trim(Str(cmbRun)))
'   iLength = 5 - iLength
'   If iLength < 0 Then iLength = 0

   On Error GoTo DiaErr1

   'bResponse = GetPartAccounts(sPartNumber, sCreditAcct, sDebitAcct)
   'cRunCost = Format((cRunOvHd + cRunMatl + cRunLabor + cRunExp), "######0.000")
   'Err = 0

'   'Check lots
'   bLots = CheckLotStatus()
'   lCOUNTER = GetLastActivity() + 1
'   sLotNumber = GetNextLotNumber()
'   lSysCount = lCOUNTER

   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   Dim mo As New ClassMO
   mo.PartNumber = sPartNumber
   mo.RunNumber = Val(cmbRun)
   lblPartial = mo.GetPreviousCompletions
   mo.CompleteMo bPartialCompletion, cComQty, _
      CDate(txtDte)
   
'   sPartNum = Trim(cmbPrt)
'   If Len(sPartNum) > 25 Then Compress (sPartNum)
'   If cComQty > 0 Then
'      sSql = "UPDATE PartTable SET PAQOH=PAQOH+" & cComQty & "," _
'             & "PALOTQTYREMAINING=PALOTQTYREMAINING+" & cComQty & " " _
'             & "WHERE PARTREF='" & sPartNumber & "'"
'      clsAdoCon.ExecuteSQL sSql
'   End If
'
'   'create ia & lot records if the quantity is > 0
'   'or if qty = 0, this is a final completion, and there are no other completions
'   'and there are no other
'   Dim rdo As ADODB.recordset
'   Dim totalMoQty As Currency
'   totalMoQty = 0
'   If cComQty = 0 Then
'      sSql = "select RUNYIELD from RunsTable" & vbCrLf _
'         & "WHERE RUNREF='" & sPartNumber & "' AND " _
'         & "RUNNO=" & Val(cmbRun) & " "
'      If GetDataSet(rdo) Then
'         totalMoQty = CCur(rdo.fields(0).Value)
'      End If
'   End If
'
'   If cComQty > 0 Or (cComQty = 0 And totalMoQty = 0) Then
'      sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
'             & "INPDATE,INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," _
'             & "INMOPART,INMORUN,INTOTMATL,INTOTLABOR,INTOTEXP," _
'             & "INTOTOH,INTOTHRS,INWIPLABACCT,INWIPMATACCT," _
'             & "INWIPOHDACCT,INWIPEXPACCT,INNUMBER,INLOTNUMBER,INUSER) " _
'             & "VALUES(6,'" & sPartNumber & "','COMPLETED RUN'," _
'             & "'RUN " & String(iLength, Chr(32)) & Trim(Str(cmbRun)) & "'," _
'             & "'" & txtDte & "','" & vDate & "'," & cComQty & "," & cComQty & "," _
'             & cRunCost & ",'" & sCreditAcct & "','" & sDebitAcct & "','" _
'             & sPartNumber & "'," & Val(cmbRun) & ","
'      sSql = sSql & cRunMatl & "," & cRunLabor & "," _
'             & cRunExp & "," & cRunOvHd & "," & cRunHours & ",'" _
'             & sInvLabAcct & "','" & sInvMatAcct & "','" _
'             & sInvExpAcct & "','" & sInvOhdAcct & "'," & lCOUNTER & ",'" _
'             & sLotNumber & "','" & sInitials & "')"
'      clsAdoCon.ExecuteSQL sSql
'
'      If bPartialCompletion Then
'         sMsg = "MO PA-"
'      Else
'         sMsg = "MO CO-"
'      End If
'      sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
'             & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
'             & "LOTUNITCOST,LOTMOPARTREF,LOTMORUNNO," _
'             & "LOTTOTMATL,LOTTOTLABOR,LOTTOTEXP,LOTTOTOH,LOTTOTHRS) " _
'             & "VALUES('" _
'             & sLotNumber & "','" & sMsg & sPartNum & " R" & Trim(Val(cmbRun)) _
'             & "','" & sPartNumber & "','" & vDate & "'," & cComQty & "," & Trim(Str(cComQty)) _
'             & "," & cRunCost & ",'" & sPartNumber & "'," _
'             & Val(cmbRun) & "," & cRunMatl & "," & cRunLabor & "," _
'             & cRunExp & "," & cRunOvHd & "," & cRunHours & ")"
'      clsAdoCon.ExecuteSQL sSql
'      If bPartialCompletion Then
'         sMsg = "MO Run Partial Comp"
'      Else
'         sMsg = "MO Run Completion"
'      End If
'      sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
'             & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
'             & "LOIMOPARTREF,LOIMORUNNO,LOIACTIVITY,LOICOMMENT) " _
'             & "VALUES('" _
'             & sLotNumber & "',1,6,'" & sPartNumber _
'             & "','" & txtDte & "'," & Trim(Str(cComQty)) _
'             & ",'" & sPartNumber & "'," & Val(cmbRun) & "," _
'             & lCOUNTER & ",'MO Run Completion')"
'      clsAdoCon.ExecuteSQL sSql
'   End If
'
'   'partial completion
'   If bPartialCompletion Then
'      Err.Clear
'      sSql = "UPDATE RunsTable SET " _
'             & "RUNYIELD=" & cComQty & "," _
'             & "RUNPARTIALQTY=RUNPARTIALQTY+" & cComQty & "," _
'             & "RUNPARTIALDATE='" & txtDte & "'," _
'             & "RUNREMAININGQTY=RUNREMAININGQTY-" & cComQty & "," _
'             & "RUNLOTNUMBER='" & sLotNumber & "' " _
'             & "WHERE RUNREF='" & sPartNumber & "' AND " _
'             & "RUNNO=" & Val(cmbRun) & " "
'      clsAdoCon.ExecuteSQL sSql
'
'   'full completion
'   Else
'      sSql = "UPDATE RnopTable SET OPCOMPDATE='" & txtDte & "'," _
'             & "OPCOMPLETE=1 WHERE OPREF='" & sPartNumber & "' AND " _
'             & "OPRUN=" & Val(cmbRun) & " AND OPCOMPLETE=0 "
'      clsAdoCon.ExecuteSQL sSql
'      sSql = "UPDATE RunsTable SET RUNCOMPLETE='" & txtDte & "'," _
'             & "RUNYIELD=" & cComQty & ",RUNSTATUS='" & sStatus & "'," _
'             & "RUNCOST=" & cRunCost & "," _
'             & "RUNOHCOST=" & cRunOvHd & "," _
'             & "RUNCMATL=" & cRunMatl & "," _
'             & "RUNCEXP=" & cRunExp & "," _
'             & "RUNCHRS=" & cRunHours & "," _
'             & "RUNCLAB=" & cRunLabor & "," _
'             & "RUNREMAININGQTY=0," _
'             & "RUNLOTNUMBER='" & sLotNumber & "' " _
'             & "WHERE RUNREF='" & sPartNumber & "' AND " _
'             & "RUNNO=" & Val(cmbRun) & " "
'      clsAdoCon.ExecuteSQL sSql
'   End If
   MouseCursor 0

'   Dim mo As New ClassMO
'   Dim sReturnMsg As String
'   sReturnMsg = mo.UpdateMOCosts(sPartNumber, Val(cmbRun), "COMPLETED RUN", _
'             cRunMatl, cRunLabor, cRunExp, cRunOvHd, cRunHours)

   clsADOCon.CommitTrans
   'AverageCost (sPartNumber)
   If CheckLotStatus <> 0 And bPartLot And cComQty > 0 Then
      MsgBox "The run was completed and lot number " & vbCrLf _
         & "System: " & mo.lotNumber & " was created." & vbCrLf _
         & "You may now edit the lot.", _
         vbInformation, Caption

'      LotEdit.optMo.Value = vbChecked
'      LotEdit.DebitAccount = sDebitAcct
'      LotEdit.CreditAccount = sDebitAcct
'      LotEdit.MONUMBER = Compress(cmbPrt)
'      LotEdit.MORUN = cmbRun
'      LotEdit.INVACTIVITY = lCOUNTER
'      LotEdit.lblOrigQty = cComQty
'
'      LotEdit.RUNCOST = cRunCost
'      LotEdit.RUNOVHD = cRunOvHd
'      LotEdit.RUNMATL = cRunMatl
'      LotEdit.RUNEXP = cRunExp
'      LotEdit.RUNHRS = cRunHours
'      LotEdit.RUNLBR = cRunLabor
'
'      LotEdit.txtLong = "MO Completion"
'      LotEdit.txtlot = "MO CO-" & cmbPrt & " Run " & Trim(Val(cmbRun))
'
'      LotEdit.INVLABACCT = sInvLabAcct
'      LotEdit.INVMATACCT = sInvMatAcct
'      LotEdit.INVEXPACCT = sInvExpAcct
'      LotEdit.INVOHDACCT = sInvOhdAcct
'
'      LotEdit.lblPart = cmbPrt
'      LotEdit.lblDate = Format(ES_SYSDATE, "mm/dd/yy")
'      LotEdit.lblTime = Format(ES_SYSDATE, "hh:mm")
      LotEdit.lblNumber = mo.lotNumber
      LotEdit.ReadExistingMoData
      LotEdit.Show 1
   End If

   'UpdateWipColumns lSysCount
   SysMsg "Manufacturing Order Is Complete.", True, Me

   On Error Resume Next
   'cmbRun.Clear
   MouseCursor 0
   cmbPrt.SetFocus

   Exit Sub

DiaErr1:
   sProcName = "CompleteMO"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   On Error Resume Next
   clsADOCon.RollbackTrans
   DoModuleErrors Me

End Sub


'Private Sub GetWipAccounts()
'   sProcName = "getlaboracct"
'   sInvLabAcct = GetLaborAcct(sPartNumber, lblCode, Val(lblLvl))
'   sProcName = "getexpenseacct"
'   sInvExpAcct = GetExpenseAcct(sPartNumber, lblCode, Val(lblLvl))
'   sProcName = "getmaterialacct"
'   sInvMatAcct = GetMaterialAcct(sPartNumber, lblCode, Val(lblLvl))
'   sProcName = "getoverheadacct"
'   sInvOhdAcct = GetOverHeadAcct(sPartNumber, lblCode, Val(lblLvl))
'End Sub
'
Private Function GetOpCompletions() As Integer
   Dim RdoOps As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT COUNT(OPCOMPLETE) from RnopTable" & vbCrLf _
      & "WHERE (OPCOMPLETE=0 AND OPREF='" & Compress(cmbPrt) & "' AND OPRUN=" & Val(cmbRun) & ") "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOps, ES_FORWARD)
   If bSqlRows Then
      With RdoOps
         If Not IsNull(.Fields(0)) Then _
                       GetOpCompletions = .Fields(0) _
                       Else GetOpCompletions = 0
         ClearResultSet RdoOps
      End With
   Else
      GetOpCompletions = 0
   End If
   
   Set RdoOps = Nothing
   
   If Err > 0 Then GetOpCompletions = 0
   
End Function

Private Function AllowOverMOQty() As Boolean
   Dim RdoCmn As ADODB.Recordset
   AllowOverMOQty = False
   On Error Resume Next
   sSql = "SELECT ISNULL(ALLOWOVERQTYCOMP, 0) ALLOWOVERQTYCOMP FROM ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmn, ES_FORWARD)
   If bSqlRows Then
      With RdoCmn
         If Val("" & !ALLOWOVERQTYCOMP) = 1 Then AllowOverMOQty = True
         ClearResultSet RdoCmn
      End With
   End If
   Set RdoCmn = Nothing
End Function

Private Function AllowMOCompletionWhenNotPC() As Boolean
   Dim RdoCmn As ADODB.Recordset
   AllowMOCompletionWhenNotPC = True
   On Error Resume Next
   sSql = "SELECT CODONTALLOWMONOTPC FROM ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmn, ES_FORWARD)
   If bSqlRows Then
      With RdoCmn
         If Val("" & !CODONTALLOWMONOTPC) = 1 Then AllowMOCompletionWhenNotPC = False
         ClearResultSet RdoCmn
      End With
   End If
   Set RdoCmn = Nothing
End Function

'Public Sub GetPreviousCompletions()
'   'update run quantities
'   Dim RdoPrev As ADODB.recordset
'   Dim cQtyIn As Currency 'Type 6
'   Dim cQtyOut As Currency 'Type 38
'   Dim cQtyBal As Currency 'Total Completed
'
'   'Complete
'   On Error GoTo DiaErr1
'   MouseCursor 13
'   sSql = "SELECT SUM(INAQTY) AS QtyComplete FROM InvaTable WHERE (INTYPE=6 " _
'          & "AND INMOPART='" & Compress(cmbPrt) & "' AND INMORUN=" & Val(cmbRun) & ")"
'   bsqlrows = clsadocon.getdataset(ssql, RdoPrev, ES_FORWARD)
'   If bSqlRows Then
'      With RdoPrev
'         If Not IsNull(!QtyComplete) Then
'            cQtyIn = !QtyComplete
'         Else
'            cQtyIn = 0
'         End If
'         ClearResultSet RdoPrev
'      End With
'   End If
'
'   If cQtyIn > 0 Then
'      sSql = "SELECT SUM(INAQTY) AS QtyComplete FROM InvaTable WHERE (INTYPE=38 " _
'             & "AND INMOPART='" & Compress(cmbPrt) & "' AND INMORUN=" & Val(cmbRun) & ")"
'      bsqlrows = clsadocon.getdataset(ssql, RdoPrev, ES_FORWARD)
'      If bSqlRows Then
'         With RdoPrev
'            If Not IsNull(!QtyComplete) Then
'               cQtyOut = !QtyComplete
'            Else
'               cQtyOut = 0
'            End If
'            ClearResultSet RdoPrev
'         End With
'      End If
'   End If
'   cQtyBal = cQtyIn - Abs(cQtyOut)
'   If cQtyBal < 0 Then cQtyBal = 0
'   sSql = "UPDATE RunsTable SET RUNPARTIALQTY=" & cQtyBal & "," & vbCrLf _
'      & "RUNREMAININGQTY=RUNQTY - " & cQtyBal & vbCrLf _
'      & "WHERE (RUNREF='" _
'      & sPartNumber & "' AND RUNNO=" & Val(cmbRun) & ")"
'   clsAdoCon.ExecuteSQL sSql
'   lblPartial = Format(cQtyIn, ES_QuantityDataFormat)
'   GoTo DiaErr2
'   Exit Sub
'
'DiaErr1:
'   Resume DiaErr2
'DiaErr2:
'   MouseCursor 0
'   Set RdoPrev = Nothing
'
'End Sub
