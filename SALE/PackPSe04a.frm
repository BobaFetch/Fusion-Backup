VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PackPSe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add An Item To A Printed Packing Slip"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSe04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5760
      TabIndex        =   27
      ToolTipText     =   "Add This Item To The Packing Slip"
      Top             =   3480
      Width           =   875
   End
   Begin VB.TextBox txtShp 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   5
      ToolTipText     =   "Amount To Ship"
      Top             =   2880
      Width           =   915
   End
   Begin VB.TextBox lblCmt 
      Enabled         =   0   'False
      Height          =   975
      Index           =   1
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "9"
      ToolTipText     =   "Comments"
      Top             =   3480
      Width           =   4215
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "PackPSe04a.frx":07AE
      DownPicture     =   "PackPSe04a.frx":1120
      Enabled         =   0   'False
      Height          =   350
      Index           =   1
      Left            =   5640
      Picture         =   "PackPSe04a.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Standard Comments"
      Top             =   3960
      Width           =   350
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Tag             =   "8"
      ToolTipText     =   "Please Select An Item From The List"
      Top             =   2160
      Width           =   615
   End
   Begin VB.ComboBox cmbItm 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Tag             =   "8"
      ToolTipText     =   "Please Select An Item From The List"
      Top             =   2160
      Width           =   615
   End
   Begin VB.ComboBox cmbSon 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Please Select An Item From The List"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5760
      TabIndex        =   4
      ToolTipText     =   "Select The Sales Order Item"
      Top             =   2160
      Width           =   875
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   240
      TabIndex        =   17
      Top             =   2040
      Width           =   6400
   End
   Begin VB.ComboBox cmbPsl 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Qualifying Packing Slips.  Select From List"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   3720
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4725
      FormDesignWidth =   6675
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ship Quantity "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5520
      TabIndex        =   26
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity/Qoh "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   25
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                             "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   24
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label lblQoh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4320
      TabIndex        =   23
      Top             =   3165
      Width           =   1035
   End
   Begin VB.Label lblOrd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      ToolTipText     =   "Ordered Quantity"
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   21
      ToolTipText     =   "Part Number"
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   20
      ToolTipText     =   "Part Description"
      Top             =   3165
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "This Function Is Intended For Packing Slips Printed, Not Shipped And Not Invoiced.  Please Select From The List."
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   15
      Top             =   960
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed"
      Height          =   255
      Index           =   23
      Left            =   3360
      TabIndex        =   14
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   1680
      Width           =   3675
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   9
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label lblPrn 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   8
      Top             =   1320
      Width           =   1035
   End
End
Attribute VB_Name = "PackPSe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'11/3/03 New
'Fixed combo
'7/8/05 Save working PS in cmbPsl
'11/17/05 Per Larry, reversed Credit/Debit accounts
'9/1/06 Correct SO Selection Query
Option Explicit
Dim bOnLoad As Byte
Dim bFIFO As Byte
Dim bComplete As Byte

Dim bLots As Byte
Dim bLotsAct As Byte
Dim cLotRemains As Currency
Dim cSalesPrice As Currency
Dim sPrintDate As String 'Preserve as printed

Dim sCurrentPs As String
Dim sCreditAcct As String
Dim sDebitAcct As String

Dim sTableDef As String
Dim sLots(50, 2) As String
'0 = Lot Number
'1 = Lot Quantity

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub CreatePsTable()
   Dim b As Byte
   Dim iRows As Integer
   Dim RdoCols As ADODB.Recordset
   Dim sTableN As String
   Dim sCol1 As Variant
   Dim sCol2 As Variant
   Dim sCol3 As Variant
   Dim sCol4 As Variant
   Dim sTable(100) As Variant
   
   MouseCursor 13
   On Error GoTo DiaErr1
   Err = 0
   clsADOCon.ADOErrNum = 0
   
   sSql = "sp_columns 'SoitTable'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCols, ES_FORWARD)
   If bSqlRows Then
      On Error GoTo 0
      With RdoCols
         sTableN = UCase(Format(Now, "ddd")) & Right(Compress(GetNextLotNumber()), 8)
         sTableDef = Trim$(sTableN)
         Do Until .EOF
            iRows = iRows + 1
            sCol1 = Trim(.Fields(3))
            sCol2 = Trim(.Fields(5))
            sCol3 = Trim(.Fields(7))
            
            sCol4 = Trim(.Fields(4))
            sCol4 = Trim(.Fields(6))
            sCol4 = Trim(.Fields(8))
            
            If sCol1 = "" Then Exit Do
            sCol4 = ""
            If iRows > 1 Then sCol1 = "," & sCol1
            If sCol2 = "char" Or sCol2 = "varchar" Then
               sCol3 = "(" & sCol3 & ") Null "
               sCol4 = "default('')"
            
            
            Else
               sCol3 = " Null "
               'BBS Added on 3/24/2010 for Ticket #15588
               'It was defaulting all the numerics to no decimal places in the temp table.
               'This was causing an implicit round. I remoevd the if statement below and replaced with this case statement
                Select Case sCol2
                    Case "smalldatetime"
                    Case "decimal"
                        sCol2 = sCol2 & "(" & Trim(.Fields(6)) & "," & Trim(.Fields(8)) & ")"
                    Case Else
                        sCol4 = "default(0)"
                End Select
'               If sCol2 <> "smalldatetime" Then sCol4 = "default(0)"
            End If
            sTable(iRows) = sCol1 & " " & sCol2 & sCol3 & sCol4
            .MoveNext
         Loop
         ClearResultSet RdoCols
         sTable(iRows) = sTable(iRows) & ")"
      End With
   End If
   If iRows > 0 And clsADOCon.ADOErrNum = 0 Then
      sTableN = "create table " & sTableDef & " ("
      For b = 1 To iRows
         sTableN = sTableN & sTable(b)
      Next
      clsADOCon.ExecuteSQL sTableN
      
      sSql = "create unique clustered index WorkRef on " & sTableDef & " " _
             & "(ITSO,ITNUMBER,ITREV) WITH FILLFACTOR=80"
      clsADOCon.ExecuteSQL sSql
   Else
      sTableDef = "PswkTable"
   End If
   Set RdoCols = Nothing
   Exit Sub
DiaErr1:
   sTableDef = "PswkTable"
   
End Sub

'5/10/02

Private Function GetPartLots(sPartWithLot As String) As Integer
   Dim RdoLots As ADODB.Recordset
   Dim iList As Integer
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
            iList = iList + 1
            sLots(iList, 0) = "" & Trim(!lotNumber)
            sLots(iList, 1) = Format$(!LOTREMAININGQTY, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoLots
      End With
      GetPartLots = iList
   Else
      GetPartLots = 0
   End If
   Set RdoLots = Nothing
   Exit Function
   
DiaErr1:
   GetPartLots = 0
   
End Function



Private Sub cmbItm_Click()
   GetItemRevisions
   
End Sub


Private Sub cmbItm_LostFocus()
   If Val(cmbItm) = 0 Then
      Beep
      cmbItm = cmbItm.List(0)
      GetItemRevisions
   End If
   
End Sub


Private Sub cmbPsl_Change()
   cmdSel.Enabled = False
   lblCmt(1).Enabled = False
   cmdComments(1).Enabled = False
   txtShp.Enabled = False
   cmdAdd.Enabled = False
   lblCmt(1).BackColor = Es_FormBackColor
   txtShp.BackColor = Es_FormBackColor
   
End Sub

Private Sub cmbPsl_Click()
   GetPackslip
   
End Sub


Private Sub cmbPsl_LostFocus()
   Dim b As Byte
   Dim iCount As Integer
   b = 0
   If Trim(cmbPsl) = "" Then _
           If cmbPsl.ListCount > 0 Then cmbPsl = cmbPsl.List(0)
   If cmbPsl.ListCount > 0 Then
      For iCount = 0 To cmbPsl.ListCount - 1
         If cmbPsl = cmbPsl.List(iCount) Then b = 1
      Next
   End If
   If b = 0 Then
      Beep
      cmbPsl = cmbPsl.List(0)
   End If
   
End Sub

Private Sub cmbRev_LostFocus()
   cmbRev = Compress(cmbRev)
   
End Sub


Private Sub cmbSon_Click()
   GetSoItems
   
End Sub


Private Sub cmbSon_LostFocus()
   Dim b As Byte
   Dim iCount As Integer
   b = 0
   If Trim(cmbSon) = "" Then _
           If cmbSon.ListCount > 0 Then cmbSon = cmbSon.List(0)
   If cmbSon.ListCount > 0 Then
      For iCount = 0 To cmbSon.ListCount - 1
         If cmbSon = cmbSon.List(iCount) Then b = 1
      Next
   End If
   If b = 0 Then
      Beep
      cmbSon = cmbSon.List(0)
   End If
   GetSoItems
   
End Sub


Private Sub cmdAdd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   bComplete = 1
   If Val(txtShp) = 0 Then
      MsgBox "Requires A Quantity Greater Than Zero.", _
         vbInformation, Caption
      txtShp = lblOrd
      Exit Sub
   End If
   If Val(txtShp) > lblOrd Then
      sMsg = "The Quantity Marked To Add Is More Than Ordered." & vbCrLf _
             & "Do You Wish Continue And Overship Anyway?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbNo Then
         CancelTrans
         Exit Sub
      End If
   End If
   If Val(txtShp) < lblOrd Then
      sMsg = "The Quantity Marked To Add Is Less Than Ordered." & vbCrLf _
             & "Do You Wish Continue And Undership Anyway?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbNo Then
         CancelTrans
         Exit Sub
      Else
         sMsg = "Is The Item Complete?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbNo Then bComplete = 0
         
         sMsg = "The Quantity To Be Added Is " & txtShp & " And Is "
         If bComplete = 0 Then
            sMsg = sMsg & vbCrLf & "Incomplete. Okay To Proceed?"
         Else
            sMsg = sMsg & vbCrLf & "Complete. Okay To Proceed?"
         End If
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbNo Then
            CancelTrans
            Exit Sub
         End If
      End If
   End If
   If bLotsAct = 1 And bLots = 1 Then
      If cLotRemains < Val(txtShp) Then
         MsgBox "This Part Is Lot Tracked An Requires Lot Selection." & vbCrLf _
            & "The Lot Quantity Remaining Won't Fill The Need.", _
            vbInformation, Caption
         Exit Sub
      End If
   End If
   sMsg = "Are You Ready To Include This Item And Mark It Printed?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then AddThisItem Else _
                  CancelTrans
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdComments_Click(Index As Integer)
   If cmdComments(1) Then
      SysComments.lblControl = "lblCmt(1)"
      SysComments.lblListIndex = 3
      SysComments.Show
      cmdComments(1) = False
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2204
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSel_Click()
   GetCurrentItem
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CreatePsTable
      bFIFO = GetInventoryMethod()
      bLotsAct = CheckLotStatus()
      FillCombo
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
   On Error Resume Next
   FormUnload
   If UCase$(Left$(sTableDef, 2)) <> "PS" Then
      sSql = "DROP TABLE " & sTableDef
      clsADOCon.ExecuteSQL sSql ' rdExecDirect
   End If
   Set PackPSe04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblInfo.ForeColor = ES_BLUE
   cmbPsl.ForeColor = ES_BLUE
   cmbSon.ForeColor = ES_BLUE
   cmbItm.ForeColor = ES_BLUE
   cmbRev.ForeColor = ES_BLUE
   cmdSel.Enabled = False
   cmdAdd.Enabled = False
   txtShp.Enabled = False
   lblCmt(1).Enabled = False
   lblCmt(1).BackColor = Es_FormBackColor
   txtShp.BackColor = Es_FormBackColor
   
End Sub

Private Sub FillCombo()
'   On Error GoTo DiaErr1
'   cmbPsl.Clear
'   lblQoh.ToolTipText = "Quantity On Hand"
'   sSql = "SELECT PSNUMBER,PSCUST,PSINVOICE FROM PshdTable WHERE " _
'          & "(PSPRINTED IS NOT NULL AND PSINVOICE=0 AND " _
'          & "PSSHIPPED=0) ORDER BY RIGHT(PSNUMBER,6) DESC"
'   LoadComboBox cmbPsl, -1
'   If cmbPsl.ListCount > 0 Then
'      If sCurrentPs = "" Then
'         cmbPsl = cmbPsl.List(0)
'      Else
'         cmbPsl = sCurrentPs
'      End If
'   Else
'      MsgBox "No Qualifying Packing Slips Found.", _
'         vbInformation, Caption
'   End If
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillcombo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
   Dim ps As New ClassPackSlip
   ps.FillPSComboPrintedNotShipped cmbPsl
      
   If cmbPsl.ListCount > 0 Then
      If sCurrentPs <> "" Then
         cmbPsl = sCurrentPs
      End If
   Else
      MsgBox "No Qualifying Packing Slips Found.", _
         vbInformation, Caption
   End If
   
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
         sPrintDate = Format(!PSPRINTED, "mm/dd/yyyy hh:mm")
         ClearResultSet RdoPsl
      End With
      FindCustomer Me, lblCst
   Else
      lblCst = ""
      lblDte = ""
      lblPrn = ""
   End If
   If lblCst <> "" Then GetSalesOrders
   Set RdoPsl = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpackslip"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetSalesOrders()
   Dim RdoSon As ADODB.Recordset
   Dim lSon As Long
   
   On Error GoTo DiaErr1
   cmdSel.Enabled = False
   lblCmt(1).Enabled = False
   cmdComments(1).Enabled = False
   txtShp.Enabled = False
   cmdAdd.Enabled = False
   lblCmt(1).BackColor = Es_FormBackColor
   txtShp.BackColor = Es_FormBackColor
   
   cmbSon.Clear
   cmbItm.Clear
   cmbRev.Clear
   sSql = "SELECT DISTINCT SONUMBER,SOCUST,ITSO FROM " _
          & "SohdTable,SoitTable WHERE SONUMBER=ITSO AND (SOCUST='" _
          & Compress(lblCst) & "' AND (ITPSITEM + ITINVOICE=0)) ORDER BY " _
          & "SONUMBER"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
   If bSqlRows Then
      With RdoSon
         Do Until .EOF
            If lSon <> !SoNumber Then _
               AddComboStr cmbSon.hWnd, Format(!SoNumber, SO_NUM_FORMAT)
            lSon = !SoNumber
            .MoveNext
         Loop
         ClearResultSet RdoSon
      End With
   End If
   If cmbSon.ListCount > 0 Then
      cmbSon = cmbSon.List(0)
      GetSoItems
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getsalesorde"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetSoItems()
   Dim RdoItm As ADODB.Recordset
   Dim iItem As Integer
   cmbItm.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT ITSO,ITNUMBER,ITACTUAL,ITCANCELED FROM SoitTable WHERE " _
          & "(ITSO=" & Val(cmbSon) & " AND ITACTUAL IS NULL AND ITCANCELED=0)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   If bSqlRows Then
      With RdoItm
         Do Until .EOF
            If iItem <> !ITNUMBER Then _
               AddComboStr cmbItm.hWnd, Format(!ITNUMBER, "##0")
            iItem = !ITNUMBER
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
   End If
   If cmbItm.ListCount > 0 Then
      cmdSel.Enabled = True
      cmbItm = cmbItm.List(0)
      GetItemRevisions
   End If
   Set RdoItm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsoitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetItemRevisions()
   cmbRev.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT ITSO,ITNUMBER,ITREV,ITCANCELED FROM SoitTable WHERE " _
          & "(ITSO=" & Val(cmbSon) & " AND ITNUMBER=" & Val(cmbItm) _
          & " AND ITACTUAL IS NULL AND ITCANCELED=0)"
   LoadComboBox cmbRev, 1
   If cmbRev.ListCount > 0 Then cmbRev = cmbRev.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "getitemrevisi"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetCurrentItem()
   Dim RdoPoi As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT ITSO,ITNUMBER,ITREV,ITQTY,ITDOLLARS,ITPART,ITCANCELED," _
          & "PARTREF,PARTNUM,PADESC,PAQOH,PALOTTRACK,PALOTQTYREMAINING " _
          & "FROM SoitTable,PartTable WHERE (ITSO=" & Val(cmbSon) & " " _
          & "AND ITNUMBER=" & Val(cmbItm) & " AND ITREV='" & Trim(cmbRev) & "' " _
          & "AND ITCANCELED=0) AND ITPART=PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPoi, ES_FORWARD)
   If bSqlRows Then
      With RdoPoi
         lblPrt(1) = "" & Trim(!PartNum)
         lblDsc(1) = "" & Trim(!PADESC)
         lblOrd = Format(!ITQTY, ES_QuantityDataFormat)
         lblQoh = Format(!PAQOH, ES_QuantityDataFormat)
         bLots = !PALOTTRACK
         cLotRemains = !PALOTQTYREMAINING
         txtShp = lblOrd
         cSalesPrice = !ITDOLLARS
         lblCmt(1).BackColor = Es_TextBackColor
         txtShp.BackColor = Es_TextBackColor
         ClearResultSet RdoPoi
      End With
      If bLotsAct Then
         lblQoh = Format(cLotRemains)
         lblQoh.ToolTipText = "Lot Quantity Available"
      Else
         lblQoh.ToolTipText = "Inventory Quantity Available"
      End If
      lblCmt(1).Enabled = True
      cmdComments(1).Enabled = True
      txtShp.Enabled = True
      cmdAdd.Enabled = True
   Else
      cLotRemains = 0
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getcurrentit"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtShp_LostFocus()
   txtShp = CheckLen(txtShp, 10)
   txtShp = Format(Abs(Val(txtShp)), ES_QuantityDataFormat)
   
End Sub



Private Sub AddThisItem()
   Dim bLotFailed As Byte
   Dim bByte As Byte
   
   Dim A As Integer
   Dim iItemNumber As Integer
   Dim iLots As Integer
   
   Dim lCOUNTER As Long
   Dim lLOTRECORD As Long
   
   Dim cItmLot As Currency
   Dim cLotQty As Currency
   Dim cPartCost As Currency
   Dim cPckQty As Currency
   Dim cQuantity As Currency
   Dim cRemPqty As Currency
   
   Dim sLot As String
   Dim sPart As String
   Dim sRev As String
   
   Dim vAdate As Variant
   
   cmdSel.Enabled = False
   cmdAdd.Enabled = False
   txtShp.Enabled = False
   lblCmt(1).Enabled = False
   MouseCursor 13
   'Okay. Do it
   'Update SoitTable
   'On Error Resume Next
   Err.Clear
   On Error GoTo whoops
   
   cQuantity = Val(txtShp)
   sPart = Compress(lblPrt(1))
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   iItemNumber = GetNextPSItem() + 1
   vAdate = Format(GetServerDateTime(), "mm/dd/yyyy hh:mm")
   sSql = "UPDATE SoitTable SET ITPSNUMBER='" & cmbPsl & "'," _
          & "ITPSITEM=" & iItemNumber & ",ITQTY=" _
          & cQuantity & " WHERE ITSO=" & Val(cmbSon) & " " _
          & "AND ITNUMBER=" & Val(cmbItm) & " " _
          & "AND ITREV='" & Trim(cmbRev) & "' "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   'Create PsitTable record
   sSql = "INSERT PsitTable (PIPACKSLIP,PIITNO,PITYPE,PIQTY,PIPART," _
          & "PISONUMBER,PISOITEM,PISOREV,PISELLPRICE,PICOMMENTS,PILOTNUMBER) " _
          & "VALUES('" & cmbPsl & "'," & iItemNumber & ",1," _
          & cQuantity & ",'" & sPart & "'," _
          & Val(cmbSon) & "," & Trim(cmbItm) & ",'" _
          & Trim(cmbRev) & "'," & cSalesPrice & ",'" _
          & lblCmt(1) & "','" & sPrintDate & "')"
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
   'is there a split required?
   If bComplete = 0 Then
      sRev = GetNextSORevision(Trim(cmbRev))
      sSql = "INSERT " & sTableDef & " SELECT * FROM SoitTable WHERE " _
             & "ITSO=" & Val(cmbSon) & " AND ITNUMBER=" & Val(cmbItm) & " AND " _
             & "ITREV='" & Trim(cmbRev) & "' "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "UPDATE " & sTableDef & " SET ITQTY=" & Val(lblOrd) - cQuantity _
             & ",ITREV='" & sRev & "' WHERE ITSO=" & Val(cmbSon) & " And " _
             & "ITNUMBER=" & Val(cmbItm) & " AND ITREV='" & Trim(cmbRev) & "' "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "INSERT SoitTable SELECT * FROM " & sTableDef & " WHERE " _
             & "ITSO=" & Val(cmbSon) & " AND ITNUMBER=" & Val(cmbItm) & " AND " _
             & "ITREV='" & sRev & "' "
     clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "UPDATE SoitTable SET ITACTUAL=NULL,ITPSNUMBER=''," _
             & "ITPSITEM=0 WHERE ITSO=" & Val(cmbSon) & " AND " _
             & "ITNUMBER=" & Val(cmbItm) & " AND ITREV='" & sRev & "' "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "DELETE FROM " & sTableDef & " WHERE " _
             & "ITSO=" & Val(cmbSon) & " AND ITNUMBER=" & Val(cmbItm) _
             & " AND ITREV='" & sRev & "' "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   End If
   
   '*** Okay now Print it (or fake print it) ***.
   'bLots = vItems(I, 5)
   cPartCost = GetPartCost(sPart, ES_STANDARDCOST)
   bByte = GetPartAccounts(lblPrt(1), sCreditAcct, sDebitAcct)
   
   'Add to Activity
   '11/17/05 Reversed Debit/Credit Accounts
   lCOUNTER = (GetLastActivity) + 1
   sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
          & "INPDATE,INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," _
          & "INPSNUMBER,INPSITEM,INNUMBER,INUSER) " _
          & "VALUES(" & IATYPE_PackingSlip & ",'" & sPart & "','PACKING SLIP'," _
          & "'" & cmbPsl & "-" & iItemNumber & "','" & vAdate & "','" _
          & vAdate & "',-" & cQuantity & ",-" & cQuantity & "," _
          & cPartCost & ",'" & sDebitAcct & "','" & sCreditAcct & "','" _
          & Trim(cmbPsl) & "'," & iItemNumber & "," & lCOUNTER & ",'" & sInitials & "')"
  clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   'Set date stamp for PsitTable 11/4/03
   sSql = "UPDATE PsitTable SET PILOTNUMBER='" & sPrintDate & "' " _
          & "WHERE PIPACKSLIP='" & Trim(cmbPsl) & "' AND PIITNO=" _
          & iItemNumber & " "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   'SO Items
   sSql = "UPDATE SoitTable SET ITACTUAL='" & vAdate _
          & "' WHERE ITPSNUMBER='" & Trim(cmbPsl) & "' AND ITNUMBER=" _
          & Val(cmbItm) & " AND ITREV='" & Trim(cmbRev) & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   'Lots
   cRemPqty = cQuantity
   If bLotsAct = 1 And bLots = 1 Then
      '***** real lots
      cLotQty = GetRemainingLotQty(sPart, True)
      If cLotQty < cQuantity Then
         MsgBox "The Lot Quantity of The Item Is Less Than Required.", _
            vbInformation, Caption
         bLotFailed = 1
      Else
         MsgBox "Lot Tracking Is Required For This Part Number." & vbCrLf _
            & "You Must Select The Lot(s) To Use For Shipping.", _
            vbInformation, Caption
         'Get The lots
         LotSelect.lblPart = sPart
         LotSelect.lblRequired = cQuantity
         LotSelect.Show vbModal
         If Es_TotalLots > 0 Then
            For A = 1 To UBound(lots)
               'insert lot transaction here
               lLOTRECORD = GetNextLotRecord(lots(A).LotSysId)
               sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                      & "LOITYPE,LOIPARTREF,LOIADATE,LOIQUANTITY," _
                      & "LOIPSNUMBER,LOIPSITEM,LOIACTIVITY,LOICOMMENT) " _
                      & "VALUES('" & lots(A).LotSysId & "'," & lLOTRECORD _
                      & "," & IATYPE_PackingSlip & ",'" & sPart & "','" & vAdate & "',-" _
                      & lots(A).LotSelQty & ",'" & cmbPsl & "'," _
                      & iItemNumber & "," & lCOUNTER & ",'Shipped Item')"
               clsADOCon.ExecuteSQL sSql 'rdExecDirect
               
               sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY" _
                      & "-" & lots(A).LotSelQty & " WHERE LOTNUMBER='" _
                      & lots(A).LotSysId & "'"
               clsADOCon.ExecuteSQL sSql 'rdExecDirect
               cItmLot = cItmLot + lots(A).LotSelQty
            Next
         Else
            MsgBox "Lot Quantity Failed And Resetting.", _
               vbInformation, Caption
            bLotFailed = 1
         End If
      End If
   Else
      iLots = GetPartLots(sPart)
      cItmLot = 0
      If iLots > 0 Then
         For A = 1 To iLots
            cLotQty = Val(sLots(A, 1))
            If cLotQty >= cRemPqty Then
               cPckQty = cRemPqty
               cLotQty = cLotQty - cRemPqty
               cRemPqty = 0
            Else
               cPckQty = cLotQty
               cRemPqty = cRemPqty - cLotQty
               cLotQty = 0
            End If
            cItmLot = cItmLot + cPckQty
            If cItmLot > Val(sLots(A, 1)) Then cItmLot = Val(sLots(A, 1))
            sLot = sLots(A, 0)
            lLOTRECORD = GetNextLotRecord(sLot)
            
            'insert lot transaction here
            sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                   & "LOITYPE,LOIPARTREF,LOIQUANTITY," _
                   & "LOIPSNUMBER,LOIPSITEM,LOICUST," _
                   & "LOIACTIVITY,LOICOMMENT) " _
                   & "VALUES('" & sLots(A, 0) & "'," _
                   & lLOTRECORD & "," & IATYPE_PackingSlip & ",'" & sPart & "',-" _
                   & Abs(cItmLot) & ",'" & Trim(cmbPsl) & "'," & iItemNumber & ",'" _
                   & Compress(lblCst) & "'," & lCOUNTER & ",'Packing Slip')"
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            
            'Update Lot Header
            sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY" _
                   & "-" & Abs(cItmLot) & " WHERE LOTNUMBER='" & sLots(A, 0) & "'"
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            If cRemPqty <= 0 Then Exit For
         Next
      End If
   End If
   If clsADOCon.ADOErrNum = 0 Then
      If bLotFailed = 1 Then
         MsgBox "The Transaction Could Not Be Completed." & vbCrLf _
            & "The Lot Quantity May Not Be Adequate.", _
            vbExclamation, Caption
         clsADOCon.RollbackTrans
      Else
         
         'update ia costs from their associated lots
         Dim ia As New ClassInventoryActivity
         ia.UpdatePackingSlipCosts (Trim(cmbPsl))
         
         'RdoCon.CommitTrans
         'Update Part Qoh and Planned date
         sSql = "UPDATE PartTable SET PAQOH=PAQOH-" & Abs(cQuantity) & " " _
                & ",PALOTQTYREMAINING=PALOTQTYREMAINING-" & Abs(cQuantity) & " " _
                & "WHERE PARTREF='" & sPart & "' "
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         
'         sSql = "UPDATE InvaTable SET INPQTY=INAQTY WHERE (INTYPE=25 " _
'                & "AND INPQTY=0 AND INPART='" & sPart & "')"
'         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         AverageCost sPart
         UpdateWipColumns lCOUNTER
         clsADOCon.CommitTrans
         
         MsgBox "The Transaction Was Completed And Marked Printed.", _
            vbInformation, Caption
      End If
   Else
      MsgBox "The Transaction Could Not Be Completed.", _
         vbExclamation, Caption
      clsADOCon.RollbackTrans
   End If
   sCurrentPs = cmbPsl
   txtShp = ""
   lblCmt(1) = ""
   lblPrt(1) = ""
   lblDsc(1) = ""
   lblOrd = ""
   lblQoh = ""
   cmbSon.Clear
   cmbItm.Clear
   cmbRev.Clear
   lblCmt(1).BackColor = Es_FormBackColor
   txtShp.BackColor = Es_FormBackColor
   FillCombo
   MouseCursor 0
   Exit Sub
   
whoops:
   clsADOCon.RollbackTrans
   MouseCursor ccHourglass
   sProcName = "AddThisItem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

'Changed to Last Item for consistency

Private Function GetNextPSItem() As Integer
   Dim RdoPsi As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT MAX(PIITNO) FROM PsitTable WHERE " _
          & "PIPACKSLIP='" & Trim(cmbPsl) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPsi, ES_FORWARD)
   If bSqlRows Then
      With RdoPsi
         If Not IsNull(.Fields(0)) Then
            GetNextPSItem = .Fields(0)
         Else
            GetNextPSItem = 1
         End If
         ClearResultSet RdoPsi
      End With
   End If
   Set RdoPsi = Nothing
   Exit Function
   
DiaErr1:
   GetNextPSItem = 1
   
End Function
