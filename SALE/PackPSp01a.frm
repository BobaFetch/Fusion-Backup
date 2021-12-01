VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form PackPSp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Packing Slips"
   ClientHeight    =   5460
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7155
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5460
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkCreateInv 
      Alignment       =   1  'Right Justify
      Caption         =   "Auto Create Invoice"
      Height          =   255
      Left            =   3360
      TabIndex        =   51
      ToolTipText     =   "This Is An Inter Company Item Transfer "
      Top             =   2640
      Width           =   1815
   End
   Begin VB.OptionButton optOrderBy 
      Caption         =   "Display Only"
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   50
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton optOrderBy 
      Caption         =   "Print Only"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   49
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox cmbEps 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "Enter Pack Slip Or Select From List (Contains Top 1000 Desc)"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CheckBox cbShowRange 
      Caption         =   "Show Range"
      Height          =   255
      Left            =   4200
      TabIndex        =   45
      Top             =   4920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox optDocList 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   43
      Top             =   4680
      Width           =   735
   End
   Begin VB.CheckBox chkPrintLabels 
      Caption         =   "Print Inventory Labels"
      Height          =   195
      Left            =   4380
      TabIndex        =   15
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "PackPSp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Show Printers"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSp01a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCpy 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "PackPSp01a.frx":0938
      Left            =   6480
      List            =   "PackPSp01a.frx":093A
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Copies To Print (Printed Only)"
      Top             =   2160
      Width           =   585
   End
   Begin VB.CheckBox optAddr 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   11
      Top             =   4920
      Width           =   735
   End
   Begin VB.CheckBox optTransfer 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Transfer"
      Height          =   255
      Left            =   5240
      TabIndex        =   14
      ToolTipText     =   "This Is An Inter Company Item Transfer "
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox optNot 
      Caption         =   "Revise Unshipped"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox optSoi 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   7
      Top             =   3720
      Width           =   735
   End
   Begin VB.CheckBox optDnp 
      Alignment       =   1  'Right Justify
      Caption         =   "Do Not Print"
      Height          =   255
      Left            =   5260
      TabIndex        =   12
      ToolTipText     =   "Adjust  Inventory, But Don't Print The Pack Slip"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CheckBox optRev 
      Caption         =   "Revise"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   4980
      TabIndex        =   34
      Top             =   420
      Width           =   2115
      Begin VB.CommandButton btnQR 
         Height          =   420
         Left            =   60
         Picture         =   "PackPSp01a.frx":093C
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Print QR Labels"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton optPrn 
         Height          =   420
         Left            =   1380
         Picture         =   "PackPSp01a.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Print The Report And Remove Items From Inventory"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton optDis 
         Height          =   420
         Left            =   720
         Picture         =   "PackPSp01a.frx":1048
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   5280
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   60
      Width           =   1365
   End
   Begin VB.CheckBox optUnp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "Fills With All Or Unprinted Pack Slips"
      Top             =   2640
      Width           =   735
   End
   Begin VB.CheckBox optFet 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optLot 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   10
      Top             =   4440
      Width           =   735
   End
   Begin VB.CheckBox optRcv 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   9
      Top             =   4200
      Width           =   735
   End
   Begin VB.CheckBox optRem 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   3960
      Width           =   735
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   3480
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   735
   End
   Begin VB.ComboBox cmbBps 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Enter Pack Slip Or Select From List (Contains Top 1000 Desc)"
      Top             =   1800
      Width           =   1215
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6420
      Top             =   3660
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5460
      FormDesignWidth =   7155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   285
      Index           =   19
      Left            =   240
      TabIndex        =   48
      Top             =   2160
      Width           =   1665
   End
   Begin VB.Label lblEpd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4200
      TabIndex        =   47
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed"
      Height          =   285
      Index           =   18
      Left            =   3480
      TabIndex        =   46
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Document List"
      Height          =   285
      Index           =   17
      Left            =   240
      TabIndex        =   44
      Top             =   4680
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PS Date"
      Height          =   285
      Index           =   16
      Left            =   240
      TabIndex        =   42
      Top             =   840
      Width           =   825
   End
   Begin VB.Label lblPSDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   41
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copies "
      Height          =   255
      Index           =   15
      Left            =   5280
      TabIndex        =   38
      ToolTipText     =   "Copies To Print (Printed Only)"
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Our Address"
      Height          =   285
      Index           =   14
      Left            =   240
      TabIndex        =   37
      Top             =   4920
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SO Item Comments"
      Height          =   285
      Index           =   13
      Left            =   240
      TabIndex        =   36
      Top             =   3720
      Width           =   1905
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   35
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Show Unprinted Only"
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   32
      Top             =   2640
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed"
      Height          =   285
      Index           =   11
      Left            =   3720
      TabIndex        =   31
      Top             =   5880
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed"
      Height          =   285
      Index           =   10
      Left            =   3480
      TabIndex        =   30
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label lblBpd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4200
      TabIndex        =   29
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Feature Options"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   28
      Top             =   5880
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Numbers"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   27
      Top             =   4440
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Receiving Validation"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   26
      Top             =   4200
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   25
      Top             =   3960
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PS Item Comments"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   24
      Top             =   3480
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   23
      Top             =   3240
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   22
      Top             =   3000
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(If Different)"
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   21
      Top             =   6120
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending PS Number"
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   20
      Top             =   5520
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PS Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   1800
      Width           =   1545
   End
End
Attribute VB_Name = "PackPSp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***

Option Explicit
Dim cmdObj1 As ADODB.Command
Dim cmdObj2 As ADODB.Command

Dim bFIFO As Byte
Dim bGoodPs1 As Byte
Dim bGoodPs2 As Byte
Dim bGoodJrn As Byte
Dim bOnLoad As Byte
Dim bPrePack As Byte

Dim iTotalItems As Integer
Dim lTransfer As Long
Dim sPackSlip As String

Dim sCustomer As String
Dim sCreditAcct As String
Dim sDebitAcct As String

Dim bGroupingByPackSlip As Boolean

Dim sLots(50, 2) As String

'0 = Lot Number
'1 = Lot Quantity

Dim vItems(800, 7) As Variant
'vItems(i , 0) = !PIPACKSLIP
'vItems(i , 1) = !PIITNO
'vItems(i , 2) = !PIQTY
' NOT USED: vItems(i , 3) = !PIPART
'vItems(i , 4) = Cost - See GetPartCost
'vItems(i , 5) = !LOTTRACK
'vItems(i , 6) = !PARTNUM

Private Const PS_PACKSLIPNO = 0
Private Const PS_ITEMNO = 1
Private Const PS_QUANTITY = 2
'Private Const PS_PIPART = 3
Private Const PS_COST = 4
Private Const PS_LOTTRACKED = 5
Private Const PS_PARTNUM = 6

Dim iUserLogo As Integer


Dim sPartGroup(800) As String '9/23/04 Compressed PartTable!PARTREF

Dim sSoItems(300, 3) As String 'Nathan 3/10/04
'0 = string of PISONUMBER
'1 = string of PISOITEM
'2 = string of PISOREV
Const SOITEM_SO = 0 ' string of PISONUMBER
Const SOITEM_ITEM = 1 ' string of PISOITEM
Const SOITEM_REV = 2 ' string of PISOREV

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'5/10/02

Private Function GetPartLots(sPartWithLot As String) As Integer
   Dim RdoLots As ADODB.Recordset
   Dim iRow As Integer
   Erase sLots
   On Error GoTo DiaErr1
   sSql = "SELECT LOTNUMBER,LOTPARTREF,LOTREMAININGQTY,LOTAVAILABLE " _
          & "FROM LohdTable WHERE (LOTPARTREF='" & sPartWithLot & "' AND " _
          & "LOTREMAININGQTY>0 AND LOTAVAILABLE=1) "
   If bFIFO = 1 Then
      sSql = sSql & "ORDER BY LOTNUMBER ASC"
   Else
      sSql = sSql & "ORDER BY LOTNUMBER DESC"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
   If bSqlRows Then
      With RdoLots
         Do Until .EOF
            If (iRow >= 49) Then Exit Do
            iRow = iRow + 1
            sLots(iRow, 0) = "" & Trim(!lotNumber)
            sLots(iRow, 1) = Format$(!LOTREMAININGQTY, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoLots
      End With
      GetPartLots = iRow
   Else
      GetPartLots = 0
   End If
   Set RdoLots = Nothing
   Exit Function
   
DiaErr1:
   GetPartLots = 0
   
End Function


Private Sub FormatControls()
   Dim b As Byte
   Dim sCustom As String
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   If ES_CUSTOM = "INTCOA" Then optTransfer.Visible = True
   sCustom = GetCustomReport("sleps01")
   If sCustom = "sleps01.rpt" Then
      z1(14).Visible = True
      optAddr.Visible = True
   End If
   For b = 1 To 8
      AddComboStr cmbCpy.hWnd, Format$(b, "0")
   Next
   AddComboStr cmbCpy.hWnd, Format$(b, "0")
   cmbCpy = cmbCpy.List(0)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   Dim ShowUnprinted As Byte
   ShowUnprinted = 0
   sOptions = GetSetting("Esi2000", "EsiSale", "Sh01", sOptions) & "000000000000"
   If Len(sOptions) > 0 Then
      optExt.Value = Val(Left(sOptions, 1))
      optCmt.Value = Val(Mid(sOptions, 2, 1))
      optRem.Value = Val(Mid(sOptions, 3, 1))
      optRcv.Value = Val(Mid(sOptions, 4, 1))
      optLot.Value = Val(Mid(sOptions, 5, 1))
      optFet.Value = Val(Mid(sOptions, 6, 1))
      ShowUnprinted = Val(Mid(sOptions, 7, 1))
      'optUnp.Value = Val(Mid(sOptions, 7, 1))
      'If Len(sOptions) = 8 Then
      optSoi.Value = Val(Mid(sOptions, 8, 1))
      optDocList.Value = Val(Mid(sOptions, 9, 1))
      chkPrintLabels.Value = Val(Mid(sOptions, 10, 1))
   End If
   
   ' trigger population of packslip dropdowns
   If ShowUnprinted = optUnp.Value Then
      FillCombo ShowUnprinted
   Else
      optUnp.Value = ShowUnprinted    ' this triggers population of lists
   End If
   
   lblPrinter = GetSetting("Esi2000", "EsiSale", "Psh01", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   optDnp.Value = GetSetting("Esi2000", "EsiSale", "Psh01dnp", optDnp.Value)
   optAddr.Value = GetSetting("Esi2000", "EsiSale", "PsAddr,", optAddr.Value)
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   On Error Resume Next
   sOptions = Trim(str(optExt.Value)) _
              & Trim(str(optCmt.Value)) _
              & Trim(str(optRem.Value)) _
              & Trim(str(optRcv.Value)) _
              & Trim(str(optLot.Value)) _
              & Trim(str(optFet.Value)) _
              & Trim(str(optUnp.Value)) _
              & Trim(str(optSoi.Value)) _
              & Trim(str(optDocList.Value)) _
              & chkPrintLabels.Value _
              & "0000000"
   SaveSetting "Esi2000", "EsiSale", "sh01", Trim(sOptions)
   SaveSetting "Esi2000", "EsiSale", "Psh01", lblPrinter
   SaveSetting "Esi2000", "EsiSale", "Psh01dnp", optDnp.Value
   SaveSetting "Esi2000", "EsiSale", "PsAddr,", optAddr.Value
   
End Sub

Private Sub btnQR_Click()

   ' make sure this report is available
   Dim emailer As New KeyMailer
   emailer.ReportName = "PSQRLabels"
   If Not emailer.GetReportInfo(True) Then
      Set emailer = Nothing
      Exit Sub
   End If

   Dim bResponse As Byte
   Dim startPS As String, endPS As String, question As String, SQL As String, msg As String
   SQL = "select PSNUMBER, PSCUST from PshdTable" & vbCrLf & "where PSNUMBER "
   startPS = cmbBps.Text
   endPS = cmbEps.Text
   If startPS = "" Then
      Exit Sub
   ElseIf endPS = "" Or endPS = startPS Then
      msg = "Create QR labels for Packing Slip " & startPS & "?"
      endPS = startPS
      SQL = SQL & "= '" & startPS & "' "
   Else
      msg = "Create QR labels for Packing Slips " & startPS & " through " & endPS & "?"
      SQL = SQL & "between '" & startPS & "' and '" & endPS & "'"
   End If
   
   ' QR labels are only for Imaginetics and only for their customer Hexcel
   SQL = SQL & " and PSCUST = 'HEXSTR'" & vbCrLf & "order by PSNUMBER"
   
   If MsgBox(msg, ES_YESQUESTION, "Print QR Labels") <> vbYes Then Exit Sub
   
   Dim rdo As ADODB.Recordset
   Dim psno As String
   Dim cust As String
   Dim labels As String
   Dim Count As Integer
   Count = 0
   msg = "QR Labels for "
   
   
   If clsADOCon.GetDataSet(SQL, rdo, ES_FORWARD) Then
      With rdo
         Do Until .EOF
            Count = Count + 1
            psno = Trim(!PsNumber)
            cust = Trim(!PSCUST)
            
            Set emailer = New KeyMailer
            emailer.ReportName = "PSQRLabels"
            If Not emailer.GetReportInfo(True) Then
               Set emailer = Nothing
               Exit Sub
            End If
            emailer.DistributionListKey = cust  ' not required for printed report
            emailer.AddStringParameter "PsNumber", psno
            If Not emailer.Generate Then
               MsgBox "Unable to create QR label"
               Set emailer = Nothing
               Return
            End If
            
            If labels <> "" Then
               labels = labels & ","
            End If
            labels = labels & psno
            .MoveNext
         Loop
      End With
   End If
   If Count = 0 Then
      MsgBox ("No Hexcel packing slips selected.  QR labels are for Hexcel only.")
   Else
      MsgBox ("QR labels for " & labels & " queued")
   End If
   rdo.Close
End Sub

Private Sub cmbBps_Click()
   If (cmbBps.Text = "MORE THAN 32767 ROWS") Then
      MsgBox "Select Valid Packslip number", vbExclamation, Caption
      Exit Sub
   End If
   
   bGoodPs1 = GetPackslip(cmbBps, lblBpd.Caption, False)
   If bGoodPs1 Then
        If optOrderBy(0) Then cmbEps = cmbBps
        lblEpd.Caption = lblBpd.Caption
   End If
End Sub

Private Sub cmbEps_LostFocus()
   If (cmbEps.Text = "MORE THAN 32767 ROWS") Then
      MsgBox "Select Valid Packslip number", vbExclamation, Caption
      Exit Sub
   End If
    
    If Len(Trim(cmbEps)) > 0 And cmbEps < cmbBps Then
        MsgBox "Must be greater than beginning Pack Slip"
        If optOrderBy(1) Then cmbEps = cmbBps
        Exit Sub
    End If
   
   
    bGoodPs2 = GetPackslip(cmbEps, lblEpd.Caption, False)
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2225
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub FillCombo(bShowUnprinted As Byte)
   Dim ps As New ClassPackSlip
   If bShowUnprinted Then
      ps.FillPSComboUnprinted cmbBps
      ps.FillPSComboUnprinted cmbEps
   Else
      ps.FillPSComboAll cmbBps
      ps.FillPSComboAll cmbEps
   End If
   
   If (cmbBps.Text = "MORE THAN 32767 ROWS") Then cmbBps.ListIndex = 0
   If (cmbEps.Text = "MORE THAN 32767 ROWS") Then cmbEps.ListIndex = 0
   
   
End Sub


Private Sub cbShowRange_Click()
   If cbShowRange.Value = vbChecked Then
      cmbEps.Visible = True
      z1(19).Caption = "PS Number To:"
      z1(19).Visible = True
      z1(0).Caption = "PS Number From:"
      lblEpd.Visible = True
      z1(18).Visible = True
   Else
      cmbEps.Visible = False
      z1(0).Caption = "PS Number"
      z1(19).Visible = False
      z1(18).Visible = False
      lblEpd.Visible = False
   End If

End Sub


Private Sub Form_Activate()
   On Error Resume Next
   If bOnLoad Then
      bPrePack = AllowPsPrepackaging()
      bFIFO = GetInventoryMethod()
      GetCompany 1
      ' Initialize show all packslips.
'      optUnp.Value = 0          ' redundant - FillCombo triggered by Form_Active setting optUnp
'
'      If optUnp.Value = 0 Then
'         FillCombo 0
'      Else
'         FillCombo 1
'      End If
      ' To check if we need to use company logo
      GetUseLogo
      'GetLastPackslip
      bOnLoad = 0
   End If
   
   If optRev.Value = vbChecked Then
      cbShowRange = vbUnchecked
      
      optOrderBy(1) = vbChecked
      'cbShowRange = vbUnchecked
      optDis.Enabled = False
      optPrn.Enabled = True
      btnQR.Enabled = True
   Else
      CheckReportGrouping
      If bGroupingByPackSlip Then cbShowRange = vbChecked Else cbShowRange = vbUnchecked
   
      optOrderBy(0) = vbChecked
      'cbShowRange = vbUnchecked
      optDis.Enabled = True
      optPrn.Enabled = False
      btnQR.Enabled = False
      
   End If
   
   'optOrderBy(1) = vbChecked

   If cbShowRange.Value = vbChecked Then
      cmbEps.Visible = True
      z1(19).Caption = "PS Number To:"
      z1(19).Visible = True
      z1(0).Caption = "PS Number From:"
      lblEpd.Visible = True
      z1(18).Visible = True
   Else
      cmbEps.Visible = False
      z1(0).Caption = "PS Number"
      z1(19).Visible = False
      z1(18).Visible = False
      lblEpd.Visible = False
   End If
   
   'Revising
   If optRev.Value = vbChecked Then
      cmbBps = PackPSe02a.cmbPsl
      cmbEps = PackPSe02a.cmbPsl
      Unload PackPSe02a
      optRev = vbUnchecked
   End If
   'Revising Unshipped
   If optNot.Value = vbChecked Then
      cmbBps = PackPSe05a.cmbPsl
      cmbEps = PackPSe05a.cmbPsl
      Unload PackPSe05a
      optNot = vbUnchecked
   End If
   
   optPrn.ToolTipText = "Print The Pack Slip And Adjust Inventory"
   optDis.ToolTipText = "Display The Pack Slip And Do Not Adjust Inventory"
   MdiSect.lblBotPanel = Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sSql = "SELECT PSCUST,PSNUMBER,PSPRINTED, PSDATE, PSBOXES FROM PshdTable WHERE PSNUMBER = ? " _
          & "AND PSTYPE=1"
   
   Set cmdObj1 = New ADODB.Command
   cmdObj1.CommandText = sSql
   'Set RdoQry1 = RdoCon.CreateQuery("", sSql)
   'RdoQry1.MaxRows = 1
   Dim prmObj1 As ADODB.Parameter
   Set prmObj1 = New ADODB.Parameter
   prmObj1.Type = adChar
   prmObj1.Size = 8
   cmdObj1.parameters.Append prmObj1
   
   sSql = "SELECT PIPACKSLIP,PIITNO,PIQTY,PIPART,PISONUMBER,PISOITEM," _
          & "PISOREV,PARTREF,PARTNUM,PALOTTRACK FROM " _
          & "PsitTable,PartTable WHERE (PIPART=PARTREF AND PIPACKSLIP = ?)" & vbCrLf _
          & "ORDER BY PIITNO"
   'Set RdoQry2 = RdoCon.CreateQuery("", sSql)
   Set cmdObj2 = New ADODB.Command
   cmdObj2.CommandText = sSql
   
   Dim prmObj2 As ADODB.Parameter
   
   Set prmObj2 = New ADODB.Parameter
   prmObj2.Type = adChar
   prmObj2.Size = 8
   cmdObj2.parameters.Append prmObj2
   
   bOnLoad = 1
   GetOptions
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set cmdObj1 = Nothing
   Set cmdObj2 = Nothing
   Set PackPSp01a = Nothing
   
End Sub

Private Sub PrintReport(Optional bPrintRange As Boolean = False)
   Dim bForm As Byte
   MouseCursor 13
   Dim sCopies As String
   
   ' if ITAR/EAR SOs, alert the user
   Dim rs As ADODB.Recordset
   Dim sos As String
   sSql = "select DISTINCT SONUMBER from PsitTable" & vbCrLf _
      & "join SohdTable on PISONUMBER = SONUMBER" & vbCrLf _
      & "where SOITAREAR = 1" & vbCrLf
   If bPrintRange Then
      sSql = sSql & "and PIPACKSLIP between '" & Trim(cmbBps) & "' and '" & Trim(cmbEps) & "'"
   Else
      sSql = sSql & "and PIPACKSLIP = '" & Trim(cmbBps) & "'"
   End If
   sSql = sSql & vbCrLf & "order by SONUMBER"
   If clsADOCon.GetDataSet(sSql, rs, ES_FORWARD) <> 0 Then
      With rs
         Do Until rs.EOF
            If Len(sos) > 0 Then
               sos = sos & ","
            End If
            sos = sos & !SoNumber
            .MoveNext
         Loop
      End With
      MsgBox "One or more SOs in ITAR/EAR status: " & sos
   End If
   rs.Close
   Set rs = Nothing
   
   
   
    On Error GoTo DiaErr1
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   'if there is a form defined, use it
   bForm = GetPrintedForm("packslip")
   If optPrn And bForm = 1 Then
      sCustomReport = GetCustomReport("slefps01")
   Else
      sCustomReport = GetCustomReport("sleps01")
   End If
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "CoAddress1"
    aFormulaName.Add "CoAddress2"
    aFormulaName.Add "CoAddress3"
    aFormulaName.Add "CoAddress4"
    aFormulaName.Add "PackSlipNumber"
    aFormulaName.Add "ShowExtComments"
    aFormulaName.Add "ShowPsComments"
    aFormulaName.Add "ShowSoItemComments"
    aFormulaName.Add "ShowRemarks"
    aFormulaName.Add "ShowReceivingValidation"
    aFormulaName.Add "ShowLotNumbers"
    aFormulaName.Add "ShowDocList"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(1)) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(2)) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(3)) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(4)) & "'")
    aFormulaValue.Add CStr("'" & CStr(sPackSlip) & "'")
    aFormulaValue.Add CStr(optExt)
    aFormulaValue.Add CStr(optCmt)
    aFormulaValue.Add CStr(optSoi)
    aFormulaValue.Add CStr(optRem)
    aFormulaValue.Add CStr(optRcv)
    aFormulaValue.Add CStr(optLot)
    aFormulaValue.Add CStr("'" & CStr(optDocList) & "'")

   
   If (iUserLogo = 1) Then
    aFormulaName.Add "ShowOurAddress"
    aFormulaValue.Add CStr(0)
   Else
    aFormulaName.Add "ShowOurAddress"
    aFormulaValue.Add CStr("'" & CStr(optAddr) & "'")
   End If
   
    aFormulaName.Add "ShowOurLogo"
    aFormulaValue.Add CStr("'" & CStr(iUserLogo) & "'")

   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.ShowGroupTree False
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
   sSql = "{PshdTable.PSNUMBER}='" & sPackSlip & "' "
   
   If bPrintRange Then
      sSql = "{PshdTable.PSNUMBER} IN '" & Trim(cmbBps) & "' TO '" & Trim(cmbEps) & "' "
   End If
   
   
   'sSql = ""
   'MdiSect.Crw.SelectionFormula = sSql
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   If optPrn Then
    cCRViewer.OpenCrystalReportObject Me, aFormulaName, Val(cmbCpy)
   Else
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
   End If
    
   
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   
   Set cCRViewer = Nothing
   
   ManageControls 0
   MouseCursor 0

   
   If PrintingKanBanLabels(sCustomer) Then
        If MsgBox("Do You Want to Print the KanBan Label(s) Now?", vbYesNoCancel, Caption) = vbYes Then
            sCopies = InputBox("Number of Copies", "KanBan Copies", 1)
            PrintKanBanLabel bPrintRange, Val(sCopies)
        End If
   End If
   
   If PrintingPaccarLabels(sCustomer) Then
        If MsgBox("Do You Want to Print the Paccar Label(s) Now?", vbYesNoCancel, Caption) = vbYes Then
            sCopies = InputBox("How many PACCAR Labels would you like to print?", "Copies", 1)
            PrintPaccarLabel bPrintRange, Val(sCopies)
        End If
   End If
   
   Exit Sub
   
DiaErr1:
   ManageControls 0
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub cmbBps_LostFocus()
'   cmbBps = CheckLen(cmbBps, 8)
'   If Val(cmbBps) > 0 Then cmbBps = "PS" & Format(cmbBps, "000000")
   If (cmbBps.Text = "MORE THAN 32767 ROWS") Then
      MsgBox "Select Valid Packslip number", vbExclamation, Caption
      Exit Sub
   End If
   
   Dim sBegPrinted As String
   bGoodPs1 = GetPackslip(cmbBps, sBegPrinted, False)
   If bGoodPs1 = 1 Then
     lblBpd.Caption = sBegPrinted
     If optOrderBy(0) Then cmbEps = cmbBps
     lblEpd.Caption = lblBpd.Caption
   End If
End Sub

Private Function GetPackslip(ByVal sPackSlipNo As String, ByRef sPrintedDate As String, DisplayMessage As Boolean, Optional ePackSlipNo As String = "") As Byte
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   

   'RdoQry1.RowsetSize = 1
   'RdoQry1(0) = Trim(cmbBps)
   cmdObj1.parameters(0).Value = Trim(sPackSlipNo)
   
   bSqlRows = clsADOCon.GetQuerySet(RdoGet, cmdObj1, ES_KEYSET, True)
   If bSqlRows Then
      With RdoGet
         sCustomer = "" & Trim(!PSCUST)
         'cmbBps = "" & Trim(!PsNumber)
         'sPackSlip = sPackSlipNo
         sPackSlip = "" & Trim(!PsNumber)
         sPrintedDate = "" & Format(!PSPRINTED, "mm/dd/yyyy")
         lblPSDate = "" & Format(!PSDATE, "mm/dd/yyyy")
         ClearResultSet RdoGet
         GetPackslip = 1
      End With
   Else
      If DisplayMessage Then
         MsgBox "Packing Slip " & sPackSlipNo & " Wasn't Found.", _
            vbInformation, Caption
      End If
      sPackSlip = ""
      GetPackslip = 0
      sCustomer = ""
      sPrintedDate = ""
   End If
   Set RdoGet = Nothing
   Exit Function
   
DiaErr1:
   GetPackslip = 0
   sProcName = "getpacksl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
Private Function EarlyLatePopEnabled() As Boolean
    Dim rdoPop As ADODB.Recordset
    Dim iDL As Integer
    
    EarlyLatePopEnabled = False
    sSql = "SELECT COPSPOPUPWARNING FROM ComnTable WHERE COREF=1"
    bSqlRows = clsADOCon.GetDataSet(sSql, rdoPop, ES_FORWARD)
    If bSqlRows Then
        iDL = 0 & rdoPop!COPSPOPUPWARNING
        If iDL = 1 Then EarlyLatePopEnabled = True
    End If
    Set rdoPop = Nothing
End Function

Private Sub SetPackSlipPrinted(Optional DontPrint As Boolean)
   
   Dim RdoPrint As ADODB.Recordset
   Dim bByte As Byte
   Dim iList As Integer
   Dim iLots As Integer
   Dim iRow As Integer
   
   Dim bInvType As Byte
   Dim bInvWritten As Byte
   Dim bLots As Byte
   Dim bLotsAct As Byte
   Dim bPrinted As Byte
   Dim bResponse As Byte
   Dim bMarkShipped As Byte
   
   Dim lSysCount As Long
   
   Dim cItmLot As Currency
   Dim cLotQty As Currency
   Dim cPartCost As Currency
   Dim cRemPqty As Currency
   Dim cPckQty As Currency
   Dim cQuantity As Currency
   
   'Costs
   Dim cMaterial As Currency
   Dim cLabor As Currency
   Dim cExpense As Currency
   Dim cOverhead As Currency
   Dim cHours As Currency
   
   Dim sMsg As String
   Dim sLot As String
   Dim sPart As String
   Dim cQtyLeft As Currency
   Dim cPsDtLotQty As Currency

   Dim Curdate As Variant
   Dim vAdate As Variant
   Dim vPSdate As Variant
   Dim vCurrentdate As Variant
   Dim bPrevMonth As Boolean
   Dim bPrevPS As Boolean
   Dim sPackSlip As String
   Dim bRet As Boolean
   
   On Error GoTo DiaErr2
   
   If Trim(cmbBps) = "" Then Exit Sub
   
   Curdate = Format(GetServerDateTime(), "mm/dd/yyyy")
   vAdate = Format(GetServerDateTime(), "mm/dd/yyyy hh:mm")
   vCurrentdate = vAdate
   vPSdate = Format(lblPSDate, "mm/dd/yyyy hh:mm")
   bPrevMonth = CheckPeriodDate(vPSdate)
   
   sPackSlip = Trim(cmbBps)
   
   If (EarlyLatePopEnabled = True) Then
      
      'getSchedule date
      Dim schdDate As String
      bRet = GetScheduleDate(sPackSlip, schdDate)
      
      If (bRet = True) Then
      
         If (CDate(lblPSDate) <> CDate(schdDate)) Then
         
            sMsg = "The PS Date is different from the Shipping Date. Do you want to print the Pack slip."
            bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
            If bResponse = vbNo Then
               Exit Sub
            End If
            ' reset the message string
            sMsg = ""
         End If
      End If
      
   Else
      If (bPrevMonth = True) Then
         sMsg = "Do you want to print the Pack slip as " & vPSdate & "."
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbNo Then
            bPrevPS = False
            vAdate = Format(GetServerDateTime(), "mm/dd/yyyy hh:mm")
         Else
            bPrevPS = True
            vAdate = Format(lblPSDate, "mm/dd/yyyy hh:mm")
         End If
         ' reset the message string
         sMsg = ""
      End If
   
   End If
   
   
   
   If optTransfer.Value = vbChecked Then bInvType = IATYPE_InvTransfer Else bInvType = IATYPE_PackingSlip

   sJournalID = GetOpenJournal("IJ", Format(vAdate, "mm/dd/yy"))
   'sJournalID = GetOpenJournal("IJ", Format(ES_SYSDATE, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      bGoodJrn = 1
   Else
      If sJournalID = "" Then bGoodJrn = 0 Else bGoodJrn = 1
   End If
   If bGoodJrn = 0 Then
      MsgBox "There Is No Open Inventory Journal For This" & vbCrLf _
         & "Period. Cannot Set The Pack Slip As Printed.", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   
   
   'Was it Printed?
   bLotsAct = CheckLotStatus()
   'RdoQry1(0) = sPackSlip
   cmdObj1.parameters(0).Value = sPackSlip
   bSqlRows = clsADOCon.GetQuerySet(RdoPrint, cmdObj1, ES_FORWARD, True)
   
   If bSqlRows Then
      With RdoPrint
         If IsNull(!PSPRINTED) Then
            bPrinted = False
         Else
            bPrinted = True
         End If
         ClearResultSet RdoPrint
      End With
   End If
   Set RdoPrint = Nothing
   'If it was printed, then print again and bail out
   If bPrinted Then
      PrintReport (True)
      Exit Sub
   End If
   
   GetItems
   If iTotalItems = 0 Then
      MsgBox "There Are No Unprinted Items On This Packing Slip.", vbInformation, Caption
      Exit Sub
   End If
   
   'quickly check that all lot-tracked items are available in sufficient quantity
   If bLotsAct Then
      For iRow = 1 To iTotalItems
         bLots = vItems(iRow, PS_LOTTRACKED)
         If bLots = 1 Then
            sPart = sPartGroup(iRow)
            cRemPqty = Val(vItems(iRow, PS_QUANTITY))
            'cLotQty = GetRemainingLotQty(sPart)
            cLotQty = GetLotRemainingQty(sPart)
            
            If (bPrevPS = True) Then
               ' Let us find if we have sufficient qty at the previous date
               cPsDtLotQty = GetLotRemainingQtyForDate(sPart, cLotQty, vAdate, vCurrentdate)
               ' Verify if the we have sufficient parts for transcations
               If (cPsDtLotQty < cRemPqty) Then
                  If sMsg = "" Then
                     sMsg = "Insufficient Lot Quantity for the following parts:" & vbCrLf
                  End If
                  sMsg = sMsg & sPart & "    required=" & cRemPqty & " available=" & cPsDtLotQty & " as of " & vAdate & vbCrLf
               End If
            Else
               If cLotQty < cRemPqty Then
                  If sMsg = "" Then
                     sMsg = "Insufficient Lot Quantity for the following parts:" & vbCrLf
                  End If
                  sMsg = sMsg & sPart & "    required=" & cRemPqty & " available=" & cLotQty & vbCrLf
               End If
            End If
         End If
      Next
      If sMsg <> "" Then
         sMsg = sMsg & "The packing slip will not be printed."
         MsgBox sMsg, vbInformation, Caption
         Exit Sub
      End If
   End If
   
   'Packing slip hasn't been printed.  Confirm that printing is desired.
   sMsg = "Do You Want To Print This Pack Slip " & vbCrLf _
          & "And Adjust Inventory For The Parts?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then Exit Sub
   
   'Custom Transfer Option
   MouseCursor 13
   If optTransfer.Value = vbUnchecked Then
      If bPrePack = 1 Then
         sMsg = "Do You Wish To Mark The Packing Slip As Shipped?" & vbCrLf _
                & "Otherwise Inventory Will Be Relieved, But The " & vbCrLf _
                & "Packing Slip Will Remain Unshipped."
         bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
         If bResponse = vbYes Then
            bMarkShipped = 1
         Else
            bMarkShipped = 0
         End If
      Else
         bMarkShipped = 1
      End If
   Else
      bResponse = MsgBox("This Is To Be Marked As A Tranfer. Continue?", _
                  ES_YESQUESTION, Caption)
      If bResponse = vbNo Then
         MouseCursor 0
         optTransfer.Value = vbUnchecked
         CancelTrans
         Exit Sub
      End If
      bMarkShipped = 1
   End If
   ManageControls 1
   
   'determine lots from which items are drawn
   sSql = "delete from TempPsLots where PsNumber = '" & sPackSlip & "'" & vbCrLf _
      & "or DateDiff( hour, WhenCreated, getdate() ) > 24"
   clsADOCon.ExecuteSql sSql ', rdExecDirect
   
   'if not using lots, start the transaction here
   If bLotsAct = 0 Then
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
   End If
   
   For iRow = 1 To iTotalItems
      sPart = sPartGroup(iRow)
      bLots = vItems(iRow, PS_LOTTRACKED)
      
      'Lots
      'cRemPqty = Val(vItems(iRow, PS_QUANTITY))
      cQuantity = Format(Val(vItems(iRow, PS_QUANTITY)), ES_QuantityDataFormat)
      If bLotsAct = 1 And bLots = 1 Then
         '***** real lots
         'cLotQty = GetRemainingLotQty(sPart)
         cLotQty = GetLotRemainingQty(sPart)
         
         If (bPrevPS = True) Then
               cPsDtLotQty = GetLotRemainingQtyForDate(sPart, cLotQty, vAdate, vCurrentdate)
               ' Verify if the we have sufficient parts for transcations
               Dim strLotMsg As String
               Dim bCheckLotTrans As Boolean
               strLotMsg = ""
               bCheckLotTrans = VerifyCurLotItQty(sPart, vAdate, vCurrentdate, cQuantity, cPsDtLotQty, strLotMsg)
               
               If (bCheckLotTrans = False) Then
                  MsgBox strLotMsg, vbInformation, Caption
                  MsgBox "Please change the Packslip printing date.", vbInformation, Caption
                  GoTo NoCanDo
               End If
               
               ' set the Lot Qty to PS printing date
               cLotQty = cPsDtLotQty
         End If
         
         If cLotQty < cQuantity Then
            MsgBox "The lot quantity of item " & vItems(iRow, PS_ITEMNO) & vbCrLf _
               & "part " & sPart & " (" & cLotQty & ") is less than required (" & cRemPqty & ")." & vbCrLf _
               & "The packing slip will not be printed."
            GoTo NoCanDo
         Else
            MsgBox "Lot tracking is required for part number " & sPart & "." & vbCrLf _
               & "You must select the lot(s) to use.", _
               vbInformation, Caption
            
            'Get The lots
            LotSelect.lblPart = vItems(iRow, PS_PARTNUM)
            LotSelect.lblRequired = Abs(cQuantity)
            LotSelect.strLotAdate = Format(vAdate, "mm/dd/yyyy")
            LotSelect.Show vbModal
            If Es_TotalLots > 0 Then
               For iList = 1 To UBound(lots)
                  'save info for this lot
                  sSql = "INSERT INTO TempPsLots ( PsNumber, PsItem, LotID, LotQty, PartRef, LotItemID)" & vbCrLf _
                     & "Values ( '" & sPackSlip & "', " & vItems(iRow, PS_ITEMNO) & ", " _
                     & "'" & lots(iList).LotSysId & "', " & lots(iList).LotSelQty _
                     & ", '" & sPart & "', '" & CStr(iList) & "') "
                  clsADOCon.ExecuteSql sSql ', rdExecDirect
               Next
            Else
               MsgBox "Lots not selected.  Packing slip will not be printed.", _
                  vbInformation, Caption
               GoTo NoCanDo
            End If
         End If
         
      'Lots off for Part or all lots
      'apply to lots automatically to the extent that lots are available
      Else
         iLots = GetPartLots(sPart)
         cItmLot = 0
         cRemPqty = Format(Val(vItems(iRow, PS_QUANTITY)), ES_QuantityDataFormat)
         
         
         If (bPrevPS = True) Then
            cPsDtLotQty = GetLotRemainingQtyForDate(sPart, cLotQty, vAdate, vCurrentdate)
            ' Verify if the we have sufficient parts for transcations
            'Dim strLotMsg As String
            strLotMsg = ""
            bCheckLotTrans = VerifyCurLotItQty(sPart, vAdate, vCurrentdate, cRemPqty, cPsDtLotQty, strLotMsg)
            
            If (bCheckLotTrans = False) Then
               MsgBox strLotMsg, vbInformation, Caption
               MsgBox "Please change the Packslip printing date.", vbInformation, Caption
               GoTo NoCanDo
            End If
            
            ' set the Lot Qty to PS printing date
            cLotQty = cPsDtLotQty
         End If
         
         For iList = 1 To iLots
            If cRemPqty <= 0 Then
               Exit For
            End If
            cLotQty = Val(sLots(iList, 1))
            If cLotQty >= cRemPqty Then
               cPckQty = cRemPqty
               cLotQty = cLotQty - cRemPqty
               cRemPqty = 0
            Else
               cPckQty = cLotQty
               cRemPqty = cRemPqty - cLotQty
               cLotQty = 0
            End If
            If cPckQty > 0 Then
               cItmLot = cItmLot + cPckQty
               If cItmLot > Val(sLots(iList, 1)) Then cItmLot = Val(sLots(iList, 1))
               sLot = sLots(iList, 0)
               sSql = "INSERT INTO TempPsLots ( PsNumber, PsItem, LotID, LotQty , PartRef, LotItemID)" & vbCrLf _
                  & "Values ( '" & sPackSlip & "', " & vItems(iRow, PS_ITEMNO) & ", " _
                  & "'" & sLot & "', " & cPckQty _
                  & ", '" & sPart & "', '" & CStr(iList) & "') "
               clsADOCon.ExecuteSql sSql ', rdExecDirect
            End If
         Next
         ' If still we have remaining Qty we need to quit
         If (cRemPqty > 0) Then
            MsgBox "Not sufficient quantity for item " & vItems(iRow, PS_ITEMNO) _
               & " part " & sPart & " available. " & vbCrLf _
               & "It is short by (" & cRemPqty & ") quantity." & vbCrLf _
               & "The packing slip will not be printed."
            GoTo NoCanDo
         End If
         
      End If
   Next
   
   
''''''''''''''''''''''''''''''''''''''

   'we have all the lots defined and there is no more user input,
   'so now ship the packing list in a single transaction
   
   'if using lots, start the transaction here
   If bLotsAct = 1 Then
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
   End If
  
   'now that we're in the transaction, make sure that all selected lot quantities are available
   Dim rdo As ADODB.Recordset
   sSql = "select count(*) as ct" & vbCrLf _
      & "from TempPsLots tmp" & vbCrLf _
      & "join LohdTable lot on tmp.LotID = lot.LotNumber" & vbCrLf _
      & "where lot.LotRemainingQty < tmp.LotQty" & vbCrLf _
      & "and PSNUMBER = '" & sPackSlip & "'"
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      If rdo!ct > 0 Then
         If bLotsAct = 1 Then
            clsADOCon.RollbackTrans
         End If
         MsgBox "Another user has allocated quantities from the lots selected.  Please try again."
         Exit Sub
      End If
   End If
   
   If bMarkShipped = 0 Then
      sSql = "UPDATE PshdTable SET PSPRINTED='" & vAdate & "'," _
             & "PSSHIPPRINT=1,PSSHIPPED=0 WHERE " _
             & "PSNUMBER='" & sPackSlip & "' AND PSTYPE=1"
      clsADOCon.ExecuteSql sSql ', rdExecDirect
   Else
      sSql = "UPDATE PshdTable SET PSPRINTED='" & vAdate & "'," _
             & "PSSHIPPRINT=1,PSSHIPPEDDATE='" & vAdate & "'," _
             & "PSSHIPPED=1 WHERE PSNUMBER='" & sPackSlip & "' " _
             & "AND PSTYPE=1"
      clsADOCon.ExecuteSql sSql ', rdExecDirect
   End If
   If clsADOCon.RowsAffected = 0 Then
      MouseCursor 0
      MsgBox "Could Not Update The Packing Slip. The Transaction " & vbCrLf _
         & "Has Been Aborted. Try Again In A Few Minutes.", _
         vbExclamation, Caption
      ManageControls 0
      clsADOCon.RollbackTrans
      Exit Sub
   End If

   'Set date stamp for all items for this packing slip
   sSql = "UPDATE PsitTable SET PILOTNUMBER='" & Format(vAdate, "mm/dd/yy hh:mm") & "' " _
      & "WHERE PIPACKSLIP='" & sPackSlip & "' "
   clsADOCon.ExecuteSql sSql ', rdExecDirect

   'Set all related SO items' ship dates
   sSql = "UPDATE SoitTable SET ITACTUAL='" & vAdate _
      & "',ITPSSHIPPED=" & bMarkShipped & " WHERE " _
      & "ITPSNUMBER='" & sPackSlip & "' "
   clsADOCon.ExecuteSql sSql ', rdExecDirect

   'get next innumber to update wip accts later
   lSysCount = GetLastActivity + 1
  
   'loop through the packing slip items
   For iRow = 1 To iTotalItems
      'CurrentLotFailed = False
      'lCOUNTER = (GetLastActivity)
      'lSysCount = lCOUNTER + 1              'do above the loop
      cQuantity = Val(vItems(iRow, PS_QUANTITY))
      sPart = sPartGroup(iRow)
      bLots = vItems(iRow, PS_LOTTRACKED)
      
      ' set the cost as standard cost
      cPartCost = GetPartCost(sPart, ES_STANDARDCOST)
      
      ' If printing PS for previous month
      If (bPrevPS = True) Then
         Dim iItemNo As Integer
         Dim strInLotNumber As String
         ' Get INLOTNUMBER
         iItemNo = CInt(Trim(vItems(iRow, PS_ITEMNO)))
         bRet = GetLotNumber(sPart, sPackSlip, iItemNo, strInLotNumber)
         
         'if the lot is a lotcharged change the cost
         If (bRet = True) Then
            GetLotPartCost sPart, strInLotNumber, vAdate, cPartCost
         End If
      End If
      
      vItems(iRow, PS_COST) = Format(cPartCost, ES_QuantityDataFormat)
      bByte = GetPartAccounts(sPart, sCreditAcct, sDebitAcct)
  
      Dim sSql1 As String
      Dim sSql2 As String
      Dim sSql3 As String

      'create inventory activities for lots for this packing slip item
      ' Fusion 5/15/2009
      sSql1 = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2, " & vbCrLf _
         & "INNUMBER,INPDATE,INADATE,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," & vbCrLf _
         & "INPSNUMBER,INPSITEM,INLOTNUMBER,INSONUMBER,INSOITEM,INSOREV) " & vbCrLf _
         & "SELECT " & bInvType & ", '" & sPart & "', 'PACKING SLIP', "
         
    sSql2 = "tmp.PsNumber + '-' + " & "cast( tmp.PsItem as varchar(5) )," & vbCrLf _
         & "(SELECT MAX(INNUMBER) as num FROM INVATABLE) +  tmp.LotItemID," & vbCrLf _
         & "'" & vAdate & "', '" & vAdate & "',  -tmp.LotQty, " _
         & cPartCost & ", '" & sDebitAcct & "', '" & sCreditAcct & "', " & vbCrLf _
         & "'" & sPackSlip & "', " & Val(vItems(iRow, PS_ITEMNO)) & ", " _
         & "tmp.LotID, " & sSoItems(iRow, SOITEM_SO) & ", "

    sSql3 = sSoItems(iRow, SOITEM_ITEM) & ", '" & sSoItems(iRow, SOITEM_REV) & "'" & vbCrLf _
         & "FROM TempPsLots tmp" & vbCrLf _
         & "JOIN PartTable pt on tmp.PARTREF = pt.PartRef" & vbCrLf _
         & "WHERE tmp.PsNumber = '" & sPackSlip & "' AND tmp.PsItem = " & Trim(vItems(iRow, PS_ITEMNO))
         
      sSql = sSql1 & sSql2 & sSql3
         
      Debug.Print sSql
      
      clsADOCon.ExecuteSql sSql ', rdExecDirect
      
      'insert lot items for this packing slip item
      sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
         & "LOITYPE,LOIPARTREF,LOIADATE,LOIQUANTITY," & vbCrLf _
         & "LOIPSNUMBER,LOIPSITEM,LOICUST,LOIACTIVITY,LOICOMMENT) " & vbCrLf _
         & "SELECT tmp.LotID, dbo.fnGetNextLotItemNumber( tmp.LotID ), " _
         & bInvType & ", '" & sPart & "', '" & vAdate & "', " & vbCrLf _
         & "-tmp.LotQty, '" & sPackSlip & "', " _
         & Val(vItems(iRow, PS_ITEMNO)) & ", '" & sCustomer & "'," _
         & "ia.INNUMBER, 'Shipped Item'" & vbCrLf _
         & "FROM TempPsLots tmp" & vbCrLf _
         & "JOIN InvaTable ia ON ia.INPSNUMBER = tmp.PsNumber AND ia.INPSITEM = tmp.PsItem" & vbCrLf _
         & "and ia.INADATE = '" & vAdate & "' and ia.INLOTNUMBER = tmp.LotID" & vbCrLf _
         & "WHERE tmp.PsNumber = '" & sPackSlip & "' AND tmp.PsItem = " & Trim(vItems(iRow, PS_ITEMNO)) & vbCrLf _
         & "ORDER BY INNUMBER desc"
      
      Debug.Print sSql
      clsADOCon.ExecuteSql sSql ', rdExecDirect
      
      'if there are quantities not covered by lots for automatic assignment, just create an ia record
      ' Fusion 5/15/2009
      
      sSql = "SELECT " & cQuantity & " + ( SELECT ISNULL( SUM( LOIQUANTITY ), 0 ) FROM LoitTable " & vbCrLf _
         & "WHERE LOIPSNUMBER = '" & sPackSlip & "' AND LOIPSITEM = " & Trim(vItems(iRow, PS_ITEMNO)) & " )"
      If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
         Dim qtyLeft As Currency
         qtyLeft = rdo.Fields(0)
         If qtyLeft > 0 Then
            sSql1 = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2, " & vbCrLf _
               & "INNUMBER,INPDATE,INADATE,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," & vbCrLf _
               & "INPSNUMBER,INPSITEM,INLOTNUMBER,INSONUMBER,INSOITEM,INSOREV) " & vbCrLf _
               & "SELECT " & bInvType & ", '" & sPart & "', 'PACKING SLIP', "
               
            sSql2 = "'" & vItems(iRow, PS_PACKSLIPNO) & Trim(vItems(iRow, PS_ITEMNO)) & "', " & vbCrLf _
               & "(SELECT MAX(INNUMBER) as num FROM INVATABLE) + 1," & vbCrLf _
               & "'" & vAdate & "', '" & vAdate & "', " & -qtyLeft & ", " _
               & cPartCost & ", '" & sDebitAcct & "', '" & sCreditAcct & "', " & vbCrLf
               
               
            sSql3 = "'" & sPackSlip & "', " & Val(vItems(iRow, PS_ITEMNO)) & ", " _
               & "'No Lot Avail', " & sSoItems(iRow, SOITEM_SO) & ", " _
               & sSoItems(iRow, SOITEM_ITEM) & ", '" & sSoItems(iRow, SOITEM_REV) & "'"
               
            sSql = sSql1 & sSql2 & sSql3
            
            Debug.Print sSql
            clsADOCon.ExecuteSql sSql ', rdExecDirect
         End If
      End If
      rdo.Close
      
      'update quantities for part
      sSql = "UPDATE PartTable SET PAQOH=PAQOH - " & cQuantity & ", " _
             & "PALOTQTYREMAINING = PALOTQTYREMAINING - " & cQuantity & vbCrLf _
             & "WHERE PARTREF='" & sPart & "' "
      clsADOCon.ExecuteSql sSql ', rdExecDirect
      AverageCost sPart
      
   
   Next
   
   'update remaining quantity in affected lots
   sSql = "UPDATE LohdTable" & vbCrLf _
      & "SET LOTREMAININGQTY = X.TOTAL" & vbCrLf _
      & "FROM LohdTable lt" & vbCrLf _
      & "JOIN (SELECT LOINUMBER, SUM(LOIQUANTITY) AS TOTAL FROM LOITTABLE GROUP BY LOINUMBER) AS X" & vbCrLf _
      & "ON X.LOINUMBER = LOTNUMBER" & vbCrLf _
      & "WHERE LOTNUMBER IN ( SELECT LotID from TempPsLots where PsNumber = '" & sPackSlip & "' )"
   clsADOCon.ExecuteSql sSql ', rdExecDirect
 
   'update ia costs from their associated lots
   Dim ia As New ClassInventoryActivity
   ia.UpdatePackingSlipCosts (sPackSlip)

   'update transfer information
   lblBpd = Format(ES_SYSDATE, "mm/dd/yyyy")
   lblBpd.Refresh
   If optTransfer.Value = vbChecked Then
      sSql = "UPDATE SoitTable SET ITINVOICE=" & lTransfer & " WHERE " _
             & "ITPSNUMBER='" & sPackSlip & "'"
      clsADOCon.ExecuteSql sSql ', rdExecDirect
      
      sSql = "UPDATE CihdTable SET INVCUST='" & sCustomer & "' " _
             & "WHERE INVNO=" & lTransfer & " "
      clsADOCon.ExecuteSql sSql ', rdExecDirect
      
      sSql = "UPDATE PshdTable SET PSINVOICE=" & lTransfer & " WHERE " _
             & "PSNUMBER='" & sPackSlip & "'"
      clsADOCon.ExecuteSql sSql ', rdExecDirect
   End If
   MouseCursor 0
   ManageControls 0
   UpdateWipColumns lSysCount
   clsADOCon.CommitTrans 'finally, commit the transaction
   SysMsg "Packing Slip Marked As Printed", True
   optTransfer.Value = vbUnchecked
   
   'print packing slip if no errors
   If Not DontPrint Then
      PrintReport
   End If
   
'>>>> INTCOA Labels
   If chkPrintLabels.Value = vbChecked Then
      If MsgBox("Print Inventory Labels?", ES_YESQUESTION, Caption) = vbYes Then
         For iRow = 1 To iTotalItems
             PackPSp01b.lblPackingSlipNumber = sPackSlip
             PackPSp01b.lblPackingSlipItem = vItems(iRow, 1)
             PackPSp01b.lblSalesOrder = Format(sSoItems(iRow, 0), SO_NUM_FORMAT)
             PackPSp01b.lblPartNumber = vItems(iRow, 6)
             PackPSp01b.lblQuantity = vItems(iRow, 2)
             PackPSp01b.Show vbModal
         Next
      End If
   End If
'>>>>  End INTCOA Label
   
   
   Exit Sub
   
DiaErr2:
   'On Error Resume Next
   
   MouseCursor 0
   ManageControls 0
   sProcName = "SetPackSlipPrinted"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   clsADOCon.RollbackTrans
   
   MsgBox "Couldn't complete inventory adjustments. " & vbCrLf _
      & "Packing list not printed.", vbExclamation, Caption
   optTransfer.Value = vbUnchecked
   Exit Sub
   
NoCanDo:
   MouseCursor 0
   ManageControls 0
   Exit Sub
End Sub

Private Sub GetItems()
   Dim RdoItm As ADODB.Recordset
   Dim iRow As Integer
   Dim bLotsAct As Byte
   Erase vItems
   Erase sSoItems
   Erase sPartGroup
   MouseCursor 13
   
   On Error GoTo DiaErr1
   bLotsAct = CheckLotStatus()
   'RdoQry2(0) = sPackSlip
   cmdObj2.parameters(0).Value = sPackSlip
   bSqlRows = clsADOCon.GetQuerySet(RdoItm, cmdObj2, ES_KEYSET, True)
   
   If bSqlRows Then
      'On Error Resume Next
      With RdoItm
         Do Until .EOF
            iRow = iRow + 1
            vItems(iRow, PS_PACKSLIPNO) = "" & Trim(!PIPACKSLIP) & "-"
            vItems(iRow, PS_ITEMNO) = Format(!PIITNO, "##0")
            vItems(iRow, PS_QUANTITY) = Format(!PIQTY, ES_QuantityDataFormat)
            'vItems(iRow, PS_PIPART) = "" & Trim(!PIPART)
            sPartGroup(iRow) = "" & Trim(!PIPART)
            vItems(iRow, PS_COST) = "0.000"
            If bLotsAct = 1 Then
               vItems(iRow, PS_LOTTRACKED) = !PALOTTRACK
            Else
               vItems(iRow, PS_LOTTRACKED) = 0
            End If
            vItems(iRow, PS_PARTNUM) = "" & Trim(!PartNum)
            sSoItems(iRow, SOITEM_SO) = str$(!PISONUMBER)
            sSoItems(iRow, SOITEM_ITEM) = str$(!PISOITEM)
            sSoItems(iRow, SOITEM_REV) = "" & Trim(!PISOREV)
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
   End If
   iTotalItems = iRow
   Set RdoItm = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub optAddr_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optCmt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   sPackSlip = Trim(cmbBps)
   If optOrderBy(0) Then cmbEps = cmbBps
   bGoodPs1 = GetPackslip(cmbBps, lblBpd.Caption, True)
   bGoodPs2 = GetPackslip(cmbEps, lblEpd.Caption, True)
   
   If bGoodPs1 = 1 And bGoodPs2 = 1 Then
        ManageControls 1
        PrintReport True
   Else
      MsgBox "Requires A Valid Packing Slip.", _
         vbInformation, Caption
   End If
   
   optDis.Enabled = True
   optPrn.Enabled = False
   btnQR.Enabled = False
   
End Sub

Private Sub optDnp_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optExt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFet_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optFet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optLot_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optLot_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optNot_Click()
   'never visible - Revise PS is open
   
End Sub

Private Sub optOrderBy_Click(Index As Integer)
   If optOrderBy(0) Then
       cbShowRange = vbUnchecked
       optDis.Enabled = True
       optPrn.Enabled = False
       btnQR.Enabled = False
   Else
       cbShowRange = vbChecked
       optDis.Enabled = False
       optPrn.Enabled = True
       btnQR.Enabled = True
   End If
End Sub


Private Sub optPrn_Click()
   'don't allow click to happen twice
   optPrn.Enabled = False

   Dim bResponse As Byte
       
'   If cmbBps <> cmbEps Then
'     MsgBox "You Cannot Print a Range of Packslips. You Must Use the View Option For a Range", vbOKOnly
'     Exit Sub
'   End If
   
   sPackSlip = Trim(cmbBps)
   'bGoodPs1 = GetPackslip(cmbBps, lblBpd.Caption, True)
   bGoodPs1 = GetPackslip(cmbBps, lblBpd.Caption, True, Trim(cmbEps))
   If bGoodPs1 = 1 Then
      If optPrn Then
         If optDnp.Value = vbChecked Then
            SetPackSlipPrinted True
         Else
            SetPackSlipPrinted False
         End If
         
         If (chkCreateInv.Value = vbChecked) Then
         
            Dim sMsg As String
            sMsg = "Do You Wish create New Invoice and Post to Journal?"
            bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
            If bResponse = vbYes Then
               Dim cAutoInvFromPs As AutoCreateInvFromPS
               Set cAutoInvFromPs = New AutoCreateInvFromPS
               cAutoInvFromPs.Init Trim(cmbBps)
               cAutoInvFromPs.AddNewInvoice
            
            End If
         End If
         
      Else
         PrintReport
      End If
   Else
      MsgBox "Requires A Valid Packing Slip.", _
         vbInformation, Caption
   End If
   
   optPrn.Enabled = True
   optDis.Enabled = False
   btnQR.Enabled = True
   
End Sub

Private Sub optRcv_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optRem_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optRem_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optRev_Click()
   'never visible - Revise PS is open
   
End Sub




Private Sub optSoi_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optUnp_Click()
   'If Not bOnLoad Then FillCombo optUnp.Value
   FillCombo optUnp.Value
End Sub

Private Sub optUnp_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optUnp_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub




Private Sub UpDown1_Change()
   
End Sub

Private Sub z1_Click(Index As Integer)
   optUnp.Value = vbChecked
   
End Sub

Private Sub GetUseLogo()
    Dim RdoLogo As ADODB.Recordset
    Dim bRows As Boolean
    ' Assumed that COMREF is 1 all the time
    sSql = "SELECT ISNULL(COLUSELOGO, 0) as COLUSELOGO FROM ComnTable WHERE COREF = 1"
    bRows = clsADOCon.GetDataSet(sSql, RdoLogo, ES_FORWARD)

    If bRows Then
        With RdoLogo
            iUserLogo = !COLUSELOGO
        End With
        'RdoLogo.Close
        ClearResultSet RdoLogo
    End If
End Sub



'12/28/04

'Private Function GetStdCosts(ListNum As Integer, Material As Currency, Labor As Currency, _
'                             Expense As Currency, OverHead As Currency, Hours As Currency) As Byte
'
'   Dim RdoStd As ADODB.Recordset
'   sSql = "SELECT PARTREF,PALEVMATL,PALEVLABOR,PALEVEXP,PALEVOH,PALEVHRS " _
'          & "FROM PartTable WHERE PARTREF='" & sPartGroup(ListNum) & "'"
'   bSqlRows = clsADOCon.GetDataSet(sSql,RdoStd, ES_FORWARD)
'   If bSqlRows Then
'      With RdoStd
'         Material = !PALEVMATL
'         Labor = !PALEVLABOR
'         Expense = !PALEVEXP
'         OverHead = !PALEVOH
'         Hours = !PALEVHRS
'         ClearResultSet RdoStd
'      End With
'   End If
'End Function
'
'

Private Function CheckPeriodDate(ByVal strPSdate As String)
   Dim strCurDate As String
   
   strPSdate = Format(strPSdate, "mm/dd/yyyy")
   strCurDate = Format(ES_SYSDATE, "mm/dd/yyyy")
   
   If (CDate(strPSdate) < CDate(strCurDate)) Then
      CheckPeriodDate = True
   Else
      CheckPeriodDate = False
   End If
   
'   Dim bThisMonth As Byte
'   Dim bTxtMonth As Byte
'   bTxtMonth = Format(strPSdate, "m")
'   bThisMonth = Format(ES_SYSDATE, "m")
'
'   If bThisMonth > bTxtMonth Then
'      CheckPeriodDate = True
'   Else
'      CheckPeriodDate = False
'   End If
   
End Function

Private Function GetLotRemainingQtyForDate(ByVal strPart As String, _
                  ByVal CurRemQty As Currency, _
                  ByVal strPSdate As String, _
                  ByVal strCurDate As String) As Currency

   Dim RdoQty As ADODB.Recordset
   Dim delRemQty As Currency
   Dim PsDateRemQty As Currency
   
   GetLotRemainingQtyForDate = 0
   
   sSql = "SELECT ISNULL(SUM(LOIQUANTITY), 0.0000) " & vbCrLf _
            & "FROM LoitTable " & vbCrLf _
            & "WHERE LOIADATE BETWEEN DATEADD(dd, 1 ,'" & strPSdate & "') " & vbCrLf _
            & " AND DATEADD(dd, 1 , '" & strCurDate & "') " & vbCrLf _
            & " AND LOIPARTREF = '" & strPart & "'"
   
   If clsADOCon.GetDataSet(sSql, RdoQty, ES_FORWARD) Then
      ' get the Delta remaining qty
      delRemQty = RdoQty.Fields(0)
      ' remaining qty at the PS date
      PsDateRemQty = CurRemQty + Abs(delRemQty)
      'return PS date remaining Qty
      GetLotRemainingQtyForDate = PsDateRemQty
   End If


End Function

Private Function VerifyCurLotItQty(ByVal sPart As String, _
                  ByVal strPSdate As String, _
                  ByVal strCurDate As String, _
                  ByVal cReqQty As Currency, _
                  ByVal cPsDtLotQty As Currency, ByRef strLotMsg As String) As Boolean

   Dim bRows As Boolean
   Dim RdoLotit As ADODB.Recordset
   Dim cRunningLotQty As Currency
   
   VerifyCurLotItQty = False
   
   sSql = "SELECT LOINUMBER, LOIRECORD, LOIPARTREF, LOIADATE, LOIQUANTITY " & vbCrLf _
            & " FROM LoitTable " & vbCrLf _
            & " WHERE LOIADATE BETWEEN DATEADD(dd, 1 ,'" & strPSdate & "') " & vbCrLf _
            & " AND DATEADD(dd, 1 , '" & strCurDate & "') " & vbCrLf _
            & " AND LOIPARTREF = '" & sPart & "' order by loiAdate"

   bRows = clsADOCon.GetDataSet(sSql, RdoLotit, ES_FORWARD)
   cRunningLotQty = cPsDtLotQty
   
   If bRows Then
      Dim strLoiNumber As String
      Dim strLotItemNum As String
      Dim strLoiAdate As String
      
      ' reduce the qty for this PS printing
      cRunningLotQty = cRunningLotQty - cReqQty
      With RdoLotit
         Do Until .EOF
            cReqQty = Trim(!LOIQUANTITY)
            strLoiNumber = Trim(!LOINUMBER)
            strLotItemNum = Trim(!LOIRECORD)
            strLoiAdate = Trim(!LOIADATE)
            
            cRunningLotQty = cRunningLotQty + cReqQty
            .MoveNext
         Loop
         ClearResultSet RdoLotit
      End With
      
      Set RdoLotit = Nothing
      
      If (cRunningLotQty < 0) Then
         
         strLotMsg = "Not sufficient lot quantity for item " & strLotItemNum & vbCrLf _
            & "part " & sPart & " from Lot Number " & strLoiNumber & " is less than required (" & cReqQty & ")." & vbCrLf _
            & "The packing slip will not be printed."
      End If
      
   End If
   
   If (strLotMsg = "") Then
      VerifyCurLotItQty = True
   End If
   
End Function

Private Function GetLotNumber(ByVal sPart As String, ByVal sPackSlip As String, _
                  ByVal iItemNo As Integer, ByRef strInLotNumber As String) As Boolean

   
   Dim RdoLotNum As ADODB.Recordset
   
   GetLotNumber = False
   sSql = "SELECT LotID FROM TempPsLots " & vbCrLf _
            & " WHERE PARTREF = '" & sPart & "' " & vbCrLf _
            & " AND PsNumber = '" & sPackSlip & "' " & vbCrLf _
            & " AND PsItem = '" & CStr(iItemNo) & "'"
   
   If clsADOCon.GetDataSet(sSql, RdoLotNum, ES_FORWARD) Then
      strInLotNumber = CStr(RdoLotNum.Fields(0))
      GetLotNumber = True
      ClearResultSet RdoLotNum
   End If
   Set RdoLotNum = Nothing
   
End Function

Private Function GetScheduleDate(ByVal sPackSlip As String, ByRef sSchDate As String) As Boolean

   
   Dim RdoSch As ADODB.Recordset
   GetScheduleDate = False
   sSql = "Select TOP 1 ITSCHED from soitTable where itpsnumber = '" & CStr(sPackSlip) & "'"
   
   If clsADOCon.GetDataSet(sSql, RdoSch, ES_FORWARD) Then
      sSchDate = Format(RdoSch.Fields(0), "mm/dd/yyyy")
      GetScheduleDate = True
      ClearResultSet RdoSch
   End If
   Set RdoSch = Nothing
   
End Function

      



Private Function GetLotPartCost(ByVal sPart As String, ByVal strInLotNumber As String, _
                        ByVal strPSdate As String, ByRef cPartCost As Currency)
   
   Dim RdoCost As ADODB.Recordset
   Dim bUseActualCost As Boolean
   Dim strLotCostedDt As String
   Dim cLotUnitCost As Currency
   Dim cStdCost As Currency
   
'   sSql = "SELECT ISNULL(PAUSEACTUALCOST, 0), LOTDATECOSTED, LOTUNITCOST, PASTDCOST " & vbCrLf _
'               & " FROM ViewLohdPartTable WHERE LOTNUMBER = '" & strInLotNumber & "'"
   sSql = "SELECT ISNULL(PAUSEACTUALCOST, 0), LOTADATE, LOTUNITCOST, PASTDCOST " & vbCrLf _
               & " FROM ViewLohdPartTable WHERE LOTNUMBER = '" & strInLotNumber & "'"
         
   If clsADOCon.GetDataSet(sSql, RdoCost, ES_FORWARD) Then
      bUseActualCost = IIf(RdoCost.Fields(0) = 1, True, False)
      If (IsNull(RdoCost.Fields(1))) Then
         bUseActualCost = False
      Else
         strLotCostedDt = Format(RdoCost.Fields(1), "mm/dd/yyyy")
      End If
      
      cLotUnitCost = RdoCost.Fields(2)
      cStdCost = RdoCost.Fields(3)
      
      ' set cost as standard cost
      cPartCost = cStdCost
      If (bUseActualCost = True) Then
         If (CDate(strLotCostedDt) < CDate(strPSdate)) Then
            cPartCost = cLotUnitCost
         End If
      End If
      
      ClearResultSet RdoCost
   End If
   Set RdoCost = Nothing
End Function

Public Sub ManageControls(DisableControl As Byte)
   If DisableControl = 1 Then
      cmbBps.Enabled = False
   Else
      cmbBps.Enabled = True
   End If
   
End Sub



Private Sub PrintKanBanLabel(ByVal bRange As Boolean, Optional ByVal NumCopies As Integer)
   Dim FormDriver  As String
   Dim FormPort    As String
   Dim FormPrinter As String
   Dim b As Byte
   Dim iCopies As Integer
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaName As New Collection
     
   On Error GoTo DiaErr1
     
   MouseCursor 13
   
   FormPrinter = GetSetting("Esi2000", "System", "PS Label Printer", "")
   If Len(RTrim(FormPrinter)) = 0 Then
        MsgBox "Please go to View | Packslip Label to setup your label printer"
        Exit Sub
   End If
   
   If Len(Trim(FormPrinter)) > 0 Then
      b = GetPrinterPort(FormPrinter, FormDriver, FormPort)
   Else
      FormPrinter = ""
      FormDriver = ""
      FormPort = ""
   End If
   If NumCopies = 0 Then iCopies = 1 Else iCopies = NumCopies
    MakeSureBoxRecordsExist Trim(cmbBps), Trim(cmbEps)
    'get custom report name if one has been defined
    sCustomReport = GetCustomReport("sleps22.rpt")
 
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = "sleps22.rpt"
    cCRViewer.ShowGroupTree False

    sSql = "{PsitTable.PIPACKSLIP}='" & Trim(cmbBps) & "' "
    If bRange Then
        sSql = "{PsitTable.PIPACKSLIP}='" & Trim(cmbBps) & "' TO '" & Trim(cmbEps) & "' "
    End If
    
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.CRViewerSize Me
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.OpenCrystalReportObject Me, aFormulaName, iCopies, FormPrinter

    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    'cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub

DiaErr1:
   sProcName = "PrintLabels"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me


End Sub


Private Sub PrintPaccarLabel(ByVal bRange As Boolean, Optional ByVal Copies As Integer)
   Dim FormDriver  As String
   Dim FormPort    As String
   Dim FormPrinter As String
   Dim b As Byte
   Dim iCopies As Integer
 
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaName As New Collection

   MouseCursor 13
   
   If Copies > 0 Then iCopies = 1 Else iCopies = Copies
   FormPrinter = GetSetting("Esi2000", "System", "PS Label Printer", "")
   If Len(RTrim(FormPrinter)) = 0 Then
        MsgBox "Please go to View | Packslip Label to setup your label printer"
        Exit Sub
   End If
   
   If Len(Trim(FormPrinter)) > 0 Then
      b = GetPrinterPort(FormPrinter, FormDriver, FormPort)
   Else
      FormPrinter = ""
      FormDriver = ""
      FormPort = ""
   End If
   
   MakeSureBoxRecordsExist Trim(cmbBps), Trim(cmbEps)
   'get custom report name if one has been defined
    sCustomReport = GetCustomReport("sleps23.rpt")
 
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = "sleps23.rpt"
    cCRViewer.ShowGroupTree False

    sSql = "{PshdTable.PSNUMBER}='" & Trim(cmbBps) & "' " ' AND {PsitTable.PIITNO} = " & ItemNo
    If bRange Then
        sSql = "{PshdTable.PSNUMBER}='" & Trim(cmbBps) & "' TO '" & Trim(cmbEps) & "' "
    End If

    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.CRViewerSize Me
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.OpenCrystalReportObject Me, aFormulaName, iCopies, FormPrinter

    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    'cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub

DiaErr1:
   sProcName = "PrintLabels"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me


End Sub


Private Sub CheckReportGrouping()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("sleps01.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   
   bGroupingByPackSlip = cCRViewer.GroupingByField("PSNUMBER")
   Set cCRViewer = Nothing

End Sub

