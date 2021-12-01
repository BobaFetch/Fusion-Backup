VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form LotsLTe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Split A Lot"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTe03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox LotComment 
      Height          =   885
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   40
      Text            =   "LotsLTe03a.frx":07AE
      Top             =   5160
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "LotsLTe03a.frx":07B2
      DownPicture     =   "LotsLTe03a.frx":1124
      Height          =   350
      Left            =   5520
      Picture         =   "LotsLTe03a.frx":1A96
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Standard Comments"
      Top             =   3840
      Width           =   350
   End
   Begin VB.ComboBox txtSplit 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6040
      TabIndex        =   6
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "C&reate"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6040
      TabIndex        =   28
      ToolTipText     =   "Create The Split"
      Top             =   3360
      Width           =   875
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "C&ancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6040
      TabIndex        =   27
      ToolTipText     =   "Cancel The Current Transacton"
      Top             =   2280
      Width           =   875
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6040
      TabIndex        =   3
      ToolTipText     =   "Select The Current Item To Split"
      Top             =   1200
      Width           =   875
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   120
      TabIndex        =   25
      Top             =   2200
      Width           =   6800
   End
   Begin VB.ComboBox cmbLot 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Lot From List And Press Select"
      Top             =   1200
      Width           =   3840
   End
   Begin VB.TextBox txtlot 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "User Lot (10 - 40) Characters"
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox txtHgt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Mat Heght"
      Top             =   5520
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtHum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      TabIndex        =   11
      Tag             =   "3"
      ToolTipText     =   "Unit Of Measure (2)"
      Top             =   5520
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   4800
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Storage Location For This Lot"
      Top             =   3360
      Width           =   675
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Height          =   1035
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Tag             =   "9"
      ToolTipText     =   "Comments (2048)"
      Top             =   3840
      Width           =   3615
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Lot Tracked Part Numbers "
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7320
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5145
      FormDesignWidth =   7005
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   285
      Index           =   18
      Left            =   240
      TabIndex        =   38
      Top             =   1870
      Width           =   1305
   End
   Begin VB.Label lblLotLoc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   37
      ToolTipText     =   "Lot Location"
      Top             =   1875
      Width           =   615
   End
   Begin VB.Label lblActCost 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   6040
      TabIndex        =   36
      ToolTipText     =   "Costed Unit Value Of this Lot"
      Top             =   1880
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Cost"
      Height          =   255
      Index           =   17
      Left            =   4900
      TabIndex        =   35
      Top             =   1880
      Width           =   1455
   End
   Begin VB.Label lblStdCost 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   6040
      TabIndex        =   34
      ToolTipText     =   "Part Number Standard Cost"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Std Cost"
      Height          =   255
      Index           =   16
      Left            =   4900
      TabIndex        =   33
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "System ID"
      Height          =   285
      Index           =   15
      Left            =   240
      TabIndex        =   32
      Top             =   2280
      Width           =   1305
   End
   Begin VB.Label lblNewSys 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   31
      ToolTipText     =   "New System Lot Number"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "stuff down here V"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Split Quantity"
      Height          =   255
      Index           =   14
      Left            =   4900
      TabIndex        =   29
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Split Comments"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   26
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   6040
      TabIndex        =   24
      ToolTipText     =   "Remaining In This Lot"
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining"
      Height          =   255
      Index           =   3
      Left            =   4900
      TabIndex        =   23
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblLotSys 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   22
      ToolTipText     =   "Existing System Lot Number"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "System ID"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   1560
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lots"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   1200
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Lot Number"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Comments"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Of Measure"
      Enabled         =   0   'False
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   17
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Cost"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   15
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   21
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   1305
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   22
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   1305
   End
End
Attribute VB_Name = "LotsLTe03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'10/3/04 New INTCOA special
'7/27/05 Added option to edit the new lot
Option Explicit
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bUnLoad As Byte

Dim cLotRemaining As Currency
Dim cStdCost As Currency
Dim cActCost As Currency
Dim cUnitCost As Currency
Dim cSplitQty As Currency

Dim sOldLot As String
Dim sOldPart As String

Dim sLots(100, 2) As String
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbLot_Click()
   If cmbLot.ListCount > 0 And cmbLot.ListIndex < 0 _
         Then cmbLot.ListIndex = 0
      lblLotSys = sLots(cmbLot.ListIndex, 0)
      GetThisLot
      
   End Sub
   
   
   Private Sub cmbLot_LostFocus()
      Dim bByte As Byte
      Dim iList As Integer
      cmbLot = CheckLen(cmbLot, 40)
      cmdSel.Enabled = True
      For iList = 0 To cmbLot.ListCount - 1
         If cmbLot = cmbLot.List(iList) Then bByte = 1
      Next
      If bByte = 0 Then
         Beep
         If cmbLot.ListCount > 0 Then cmbLot = cmbLot.List(0)
      End If
      If cmbLot.ListCount > 0 And cmbLot.ListIndex < 0 _
            Then cmbLot.ListIndex = 0
         If cmbLot.ListCount > 0 Then
            lblLotSys = sLots(cmbLot.ListIndex, 0)
            GetThisLot
         End If
         
      End Sub
      
      
      Private Sub cmbPrt_Click()
         GetSplitPart
         
      End Sub
      
      
      Private Sub cmbPrt_LostFocus()
         cmbPrt = CheckLen(cmbPrt, 30)
         If bCancel = 1 Then Exit Sub
         
         If (Not ValidPartNumber(cmbPrt.Text)) Then
            MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
               vbInformation, Caption
            cmbPrt = ""
            Exit Sub
         End If
         
         If sOldPart <> cmbPrt Then GetSplitPart
         
      End Sub
      
      
      Private Sub cmdCan_Click()
         Unload Me
         
      End Sub
      
      
      
      Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
         bCancel = 1
         
      End Sub
      
      
      Private Sub cmdComments_Click()
         If cmdComments Then
            'See List For Index
            txtCmt.SetFocus
            SysComments.lblListIndex = 3
            SysComments.Show
            cmdComments = False
         End If
         
      End Sub
      
      Private Sub cmdEnd_Click()
         Dim bResponse As Byte
         bResponse = MsgBox("Cancel The Creation Of The Lot Split.", _
                     ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            cmdSel.Enabled = True
            cmbPrt.Enabled = True
            cmbLot.Enabled = True
            ManageBoxes 0
            cmdSplit.Enabled = False
            cmdEnd.Enabled = False
         End If
         
      End Sub
      
      Private Sub cmdHlp_Click()
         If cmdHlp Then
            MouseCursor 13
            OpenHelpContext "5503"
            MouseCursor 0
            cmdHlp = False
         End If
         
      End Sub
      
      Private Sub cmdSel_Click()
         ES_SYSDATE = Format(GetServerDateTime, "mm/dd/yy hh:mm:ss")
         ManageBoxes 1
         cLotRemaining = Format(Val(lblRem), "#########0.000")
         cmbPrt.Enabled = False
         cmbLot.Enabled = False
         cmdSel.Enabled = False
         cmdEnd.Enabled = True
         cmdSplit.Enabled = True
         lblNewSys = GetNextLotNumber()
         ' txtlot = "SPLIT-" & lblLotSys & "-" & Format(ES_SYSDATE, "yy-mm-ddhhmmss")
         txtlot = cmbLot & "A"
         sOldLot = txtlot
         txtCmt = LotComment
         If cActCost > 0 Then
            txtCst = Format(cActCost, ES_QuantityDataFormat)
            txtCst.Enabled = False
         End If
         
      End Sub
      
      
      
      Private Sub cmdSplit_Click()
         Dim bLen As Byte
         Dim sMsg As String
         
         cUnitCost = Format(Abs(Val(txtCst)), "######0.000")
         cSplitQty = Format(Abs(Val(txtQty)), "######0.000")
         If cSplitQty = 0 Then
            MsgBox "The Split Quantity Must Be Greater Than Zero.", _
               vbInformation, Caption
            Exit Sub
         End If
         bLen = Len(Trim(txtSplit))
         If bLen < 5 Or bLen > 20 Then
            MsgBox "Split Comments Between 5 And 20 Chars.", _
               vbInformation, Caption
            Exit Sub
         End If
         bLen = Len(Trim(txtLoc))
         If bLen < 1 Then
            MsgBox "Splits Require A Location.", _
               vbInformation, Caption
            Exit Sub
         End If
         bLen = Len(Trim(txtlot))
         If bLen < 5 Or bLen > 40 Then
            MsgBox "User Lot Must Be Between 5 And 40 Chars.", _
               vbInformation, Caption
            Exit Sub
         End If
         If cSplitQty >= cLotRemaining Then
            MsgBox "The Split Quantity Must Be Less Than Lot Remaining.", _
               vbInformation, Caption
            Exit Sub
         End If
         'Finished Testing.  Create It
         sMsg = "The Split Will Leave a Remainder Of " & (cLotRemaining - cSplitQty) & " In " & vbCr _
                & "The Existing Lot. Continue To Create The Split?"
         bLen = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bLen = vbYes Then CreateSplit Else CancelTrans
         
      End Sub
      
      Private Sub Form_Activate()
         Dim b As Byte
         MdiSect.lblBotPanel = Caption
         If bOnLoad Then
            ManageBoxes 0
            bUnLoad = 1
            b = CheckLotStatus()
            If b = 1 Then
               FillCombo
               bOnLoad = 0
            Else
               MsgBox "Requires Lots Be Turned On.", _
                  vbInformation, Caption
               Unload Me
            End If
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
         If bUnLoad = 1 Then FormUnload
         Set LotsLTe03a = Nothing
         
      End Sub
      
      
      
      Private Sub FormatControls()
         Dim b As Byte
         b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
         
      End Sub
      
      Private Sub FillCombo()
         On Error GoTo DiaErr1
         sOldPart = ""
         cmbPrt.Clear
         cmbLot.Clear
         sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE " _
                & "PALOTTRACK =1 ORDER BY PARTREF"
         LoadComboBox cmbPrt
         
         ' The ListCount is an Integer (only 32,000)...if the rows
         ' more than that Listcount is negative.
         If cmbPrt.ListCount <> 0 Then
            AddComboStr txtSplit.hWnd, " "
            sSql = "SELECT DISTINCT LOTSPLITCOMMENT FROM LohdTable " _
                   & "WHERE LOTSPLITCOMMENT<>'' ORDER BY LOTSPLITCOMMENT"
            LoadComboBox txtSplit, -1
            cmbPrt = cmbPrt.List(0)
            GetSplitPart
         Else
            MsgBox "No Lot Tracked Parts Were Found.", _
               vbInformation, Caption
            Unload Me
         End If
         Exit Sub
         
DiaErr1:
         sProcName = "fillcombo"
         CurrError.Number = Err.Number
         CurrError.Description = Err.Description
         DoModuleErrors Me
         
      End Sub
      
      
      Private Sub GetSplitPart()
         Dim RdoPrt As ADODB.Recordset
         cmbLot.Clear
         sOldPart = cmbPrt
         ManageBoxes 0
         cmdSel.Enabled = False
         On Error GoTo DiaErr1
         sSql = "SELECT PARTREF,PARTNUM,PADESC,PASTDCOST,PALOTQTYREMAINING " _
                & "FROM PartTable WHERE PARTREF='" & Compress(cmbPrt) & "' " _
                & "AND PALOTTRACK=1"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
         If bSqlRows Then
            With RdoPrt
               cmbPrt = "" & Trim(!PartNum)
               lblDsc = "" & Trim(!PADESC)
               lblStdCost = Format(!PASTDCOST, "######0.000")
               lblRem = Format(!PALOTQTYREMAINING, "######0.000")
               ClearResultSet RdoPrt
            End With
         Else
            lblDsc = "*** Lot Tracked Part Number Not Found ***"
         End If
         Set RdoPrt = Nothing
         If bSqlRows Then GetSplitLots
         Exit Sub
         
DiaErr1:
         sProcName = "getsplitpart"
         CurrError.Number = Err.Number
         CurrError.Description = Err.Description
         DoModuleErrors Me
         
      End Sub
      
      Private Sub lblActCost_Click()
         cActCost = Format(Val(lblActCost), ES_QuantityDataFormat)
         
      End Sub
      
      Private Sub lblDsc_Change()
         If Left(lblDsc, 6) = "*** Lo" Then _
                 lblDsc.ForeColor = ES_RED Else _
                 lblDsc.ForeColor = vbBlack
         
      End Sub
      
      Private Sub GetSplitLots()
         Dim RdoLot As ADODB.Recordset
         Dim iRow As Integer
         On Error GoTo DiaErr1
         ManageBoxes 0
         Erase lots
         iRow = -1
         sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTPARTREF,LOTREMAININGQTY FROM " _
                & "LohdTable WHERE (LOTPARTREF='" & Compress(cmbPrt) & "' " _
                & "AND LOTREMAININGQTY>0 AND LOTSPLITFROMSYS='') ORDER BY LOTUSERLOTID "
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
         If bSqlRows Then
            With RdoLot
               Do Until .EOF
                  AddComboStr cmbLot.hWnd, "" & Trim(!LOTUSERLOTID)
                  iRow = iRow + 1
                  sLots(iRow, 0) = "" & Trim(!lotNumber)
                  sLots(iRow, 1) = "" & Trim(!LOTUSERLOTID)
                  .MoveNext
               Loop
               ClearResultSet RdoLot
            End With
         End If
         
         If cmbLot.ListCount > 0 Then
            cmbLot = cmbLot.List(0)
            cmbLot.ListIndex = 0
            lblLotSys = sLots(0, 0)
            GetThisLot
            cmdSel.Enabled = True
         Else
            lblLotSys = "No Lots With Quantities Found"
         End If
         Set RdoLot = Nothing
         If bSqlRows Then GetThisLot
         Exit Sub
         
DiaErr1:
         sProcName = "getsplitlot"
         CurrError.Number = Err.Number
         CurrError.Description = Err.Description
         DoModuleErrors Me
         
      End Sub
      
      Private Sub GetThisLot()
         Dim RdoLot As ADODB.Recordset
         On Error GoTo DiaErr1
         sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTPARTREF,LOTUNITCOST," _
                & "LOTLOCATION,LOTREMAININGQTY,LOTCOMMENTS FROM LohdTable " _
                & "WHERE (LOTPARTREF='" & Compress(cmbPrt) & "' AND LOTNUMBER='" _
                & lblLotSys & "')"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
         If bSqlRows Then
            With RdoLot
               lblRem = Format(!LOTREMAININGQTY, ES_QuantityDataFormat)
               lblActCost = Format(!LotUnitCost, ES_QuantityDataFormat)
               lblLotSys = "" & Trim(!lotNumber)
               lblLotLoc = "" & Trim(!LOTLOCATION)
               LotComment = "" & Trim(!LOTCOMMENTS)
               ClearResultSet RdoLot
            End With
            cmdSel.Enabled = True
         Else
            ManageBoxes 0
            cmdSel.Enabled = False
            lblLotSys = "No Lots With Quantities Found"
            lblRem = "0.000"
         End If
         Set RdoLot = Nothing
         Exit Sub
         
DiaErr1:
         sProcName = "getthislot"
         CurrError.Number = Err.Number
         CurrError.Description = Err.Description
         DoModuleErrors Me
         
      End Sub
      
      Private Sub ManageBoxes(bEnable As Byte)
         If bEnable = 0 Then
            cmdComments.Enabled = False
            txtlot.Enabled = False
            txtSplit.Enabled = False
            txtCst.Enabled = False
            txtLoc.Enabled = False
            txtCmt.Enabled = False
            txtQty.Enabled = False
            cmdEnd.Enabled = False
            txtQty = "0.000"
            txtCst = "0.000"
            txtLoc = ""
            txtCmt = ""
            txtlot = ""
            lblNewSys = ""
            lblActCost = "0.000"
         Else
            cmdComments.Enabled = True
            txtlot.Enabled = True
            txtSplit.Enabled = True
            txtCst.Enabled = True
            txtLoc.Enabled = True
            txtCmt.Enabled = True
            txtQty.Enabled = True
         End If
         
      End Sub
      
      Private Sub lblStdCost_Click()
         cStdCost = Format(Val(lblStdCost), ES_QuantityDataFormat)
      End Sub
      
      
      Private Sub txtCmt_LostFocus()
         txtCmt = CheckLen(txtCmt, 2048)
         
      End Sub
      
      
      Private Sub txtCst_LostFocus()
         txtCst = Format(Abs(Val(txtCst)), ES_QuantityDataFormat)
         
      End Sub
      
      
      Private Sub txtLoc_LostFocus()
         txtLoc = CheckLen(txtLoc, 4)
         
      End Sub
      
      
      Private Sub txtlot_LostFocus()
         Dim bByte As Byte
         txtlot = CheckLen(txtlot, 40)
         If Len(Trim(txtlot)) < 5 Then
            Beep
            txtlot = sOldLot
            MsgBox "Requires At Least (5 chars).", _
               vbInformation
         Else
            bByte = GetUserLotID(Trim(txtlot))
            If bByte = 1 Then txtlot = sOldLot
         End If
         sOldLot = txtlot
         
      End Sub
      
      
      Private Sub txtQty_LostFocus()
         txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
         
      End Sub
      
      
      Private Sub txtSplit_LostFocus()
         Dim bByte As Byte
         Dim iList As Integer
         txtSplit = CheckLen(txtSplit, 20)
         If txtSplit.ListCount > 0 Then
            For iList = 0 To txtSplit.ListCount - 1
               If txtSplit = txtSplit.List(iList) Then bByte = 1
            Next
         End If
         If bByte = 0 Then txtSplit.AddItem txtSplit
         
      End Sub
      
      
      
      Public Sub CreateSplit()
         Dim bResponse As Byte
         
         Dim lCOUNTER As Long
         Dim lLOTRECORD As Long
         Dim lSysCount As Long
         
         Dim sLotNum As String
         Dim sPartNumber As String
         Dim vAdate As Variant
         sPartNumber = Compress(cmbPrt)
         sLotNum = Trim(lblNewSys)
         lCOUNTER = GetLastActivity()
         lSysCount = lCOUNTER + 1
         vAdate = Format(GetServerDateTime(), "mm/dd/yy hh:mm")
         lLOTRECORD = GetNextLotRecord(Trim(lblLotSys))
         'new split lot
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         clsADOCon.BeginTrans
         
         lCOUNTER = lCOUNTER + 1
         sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
                & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
                & "LOTUNITCOST,LOTCOMMENTS,LOTSPLITFROMSYS,LOTSPLITFROMUSER," _
                & "LOTSPLITCOMMENT,LOTLOCATION) VALUES('" _
                & sLotNum & "','" & txtlot & "','" & sPartNumber _
                & "','" & vAdate & "'," & cSplitQty & "," & cSplitQty _
                & "," & cUnitCost & ",'" & Trim(txtCmt) & "','" _
                & Trim(lblLotSys) & "','" & Trim(txtlot) & "','" _
                & Trim(txtSplit) & "','" & txtLoc & "')"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
                & "LOIACTIVITY,LOICOMMENT) " _
                & "VALUES('" _
                & sLotNum & "',1,19,'" & sPartNumber _
                & "','" & vAdate & "'," & cSplitQty _
                & "," & lCOUNTER & "," _
                & "'From Split Lot')"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
                & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INLOTNUMBER," _
                & "INUSER) VALUES(19,'" & sPartNumber & "','FROM SPLIT ','" & Trim(lblLotSys) & "'," _
                & "'" & vAdate & "','" & vAdate & "'," & cSplitQty _
                & "," & cSplitQty & "," & cUnitCost & ",'',''," & lCOUNTER & ",'" _
                & Trim(lblNewSys) & "','" & sInitials & "')"
         clsADOCon.ExecuteSQL sSql
         
         'Update Old Lot
         lCOUNTER = lCOUNTER + 1
         sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY-" _
                & cSplitQty & " WHERE LOTNUMBER='" & Trim(lblLotSys) & "'"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
                & "LOIACTIVITY,LOICOMMENT) " _
                & "VALUES('" _
                & Trim(lblLotSys) & "'," & lLOTRECORD & ",19,'" & sPartNumber _
                & "','" & vAdate & "',-" & cSplitQty _
                & "," & lCOUNTER & "," _
                & "'To Split Lot')"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
                & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INLOTNUMBER," _
                & "INUSER) VALUES(19,'" & sPartNumber & "','TO SPLIT ','" & Trim(lblNewSys) & "'," _
                & "'" & vAdate & "','" & vAdate & "',-" & cSplitQty _
                & ",-" & cSplitQty & "," & cUnitCost & ",'',''," & lCOUNTER & ",'" _
                & Trim(lblLotSys) & "','" & sInitials & "')"
         clsADOCon.ExecuteSQL sSql
         
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            sSql = "INSERT INTO LospTable (SPLIT_FROMLOT," _
                   & "SPLIT_TOLOT,SPLIT_QUANTITY) VALUES('" _
                   & Trim(lblLotSys) & "','" & Trim(lblNewSys) & "'," _
                   & cSplitQty & ")"
            clsADOCon.ExecuteSQL sSql
            UpdateWipColumns lSysCount
            
            SysMsg "The Lot Was Successfully Split.", True
            bResponse = MsgBox("Do You Wish To Edit The New Split Lot?", _
                        ES_NOQUESTION, Caption)
            If bResponse = vbYes Then
               bUnLoad = 0
               LotsLTe03b.cmbPrt = cmbPrt
               LotsLTe03b.lblDsc = lblDsc
               LotsLTe03b.lblNumber = lblNewSys
               LotsLTe03b.GetCalledLot
               LotsLTe03b.Show
               Unload Me
            Else
               cmbPrt.Enabled = True
               cmbLot.Enabled = True
               ManageBoxes 0
               cmdSplit.Enabled = False
               cmdEnd.Enabled = False
               cmdSel.Enabled = True
               FillCombo
            End If
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            
            MsgBox "The Lot Could Not Be Successfully Split.", _
               vbInformation, Caption
         End If
      End Sub
