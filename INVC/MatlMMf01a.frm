VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form MatlMMf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adjust Part Quantity"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z3 
      Height          =   30
      Left            =   120
      TabIndex        =   55
      Top             =   1680
      Width           =   7095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "MatlMMf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPrt1 
      Height          =   285
      Left            =   1320
      TabIndex        =   53
      Tag             =   "3"
      ToolTipText     =   "Leading Char Search  (*  In Front Is A Legal Wild Card)"
      Top             =   0
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CommandButton cmdFnd 
      Height          =   315
      Left            =   4440
      Picture         =   "MatlMMf01a.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   360
      Width           =   350
   End
   Begin VB.ComboBox cmbLot 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   2520
      Width           =   4005
   End
   Begin VB.CheckBox optLot 
      Caption         =   "Lot Tracked Part"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3000
      TabIndex        =   45
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtAvg 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Update The Standard Cost"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      TabIndex        =   14
      ToolTipText     =   "Apply Inventory Adjustment"
      Top             =   1800
      Width           =   875
   End
   Begin VB.TextBox txtHrs 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txtOvh 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   12
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtExp 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtMat 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   10
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtLab 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtStd 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "Standard Cost (Not Available To Edit)"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Tag             =   "4"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtCmt 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Tag             =   "2"
      ToolTipText     =   "Comment. Required"
      Top             =   3240
      Width           =   4005
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Tag             =   "1"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Tag             =   "1"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox optRep 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   1060
      Width           =   735
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Select Part Number From List"
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6360
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   5880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6105
      FormDesignWidth =   7380
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Qty Available"
      Height          =   255
      Index           =   23
      Left            =   3600
      TabIndex        =   58
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblExpDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   57
      ToolTipText     =   "Actual Lot Creation Date"
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblExpDateLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Expiration Date"
      Height          =   255
      Left            =   4680
      TabIndex        =   56
      Top             =   2940
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label LBLLotLoc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   51
      ToolTipText     =   "Lot Location"
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   50
      ToolTipText     =   "Actual Lot Creation Date"
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Created"
      Height          =   255
      Index           =   22
      Left            =   4080
      TabIndex        =   49
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label lblNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   48
      ToolTipText     =   "System Produced Lot Number Click To Set User Lot The Same"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "System Lot Number"
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   47
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblLotQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5040
      TabIndex        =   46
      ToolTipText     =   "Remaing Lot Qauntity"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lots QOH"
      Height          =   255
      Index           =   20
      Left            =   5160
      TabIndex        =   44
      ToolTipText     =   "Total Of Lots"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label txtLqoh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   43
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Average Cost"
      Height          =   255
      Index           =   19
      Left            =   3480
      TabIndex        =   42
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Prod Code"
      Height          =   255
      Index           =   18
      Left            =   5160
      TabIndex        =   40
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   39
      Top             =   600
      Width           =   945
   End
   Begin VB.Label lblActDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3960
      TabIndex        =   38
      Top             =   4320
      Width           =   3345
   End
   Begin VB.Label lblQoh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   37
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty On Hand"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   36
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   35
      Top             =   1800
      Width           =   350
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
      Height          =   255
      Index           =   16
      Left            =   600
      TabIndex        =   34
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Expense"
      Height          =   255
      Index           =   15
      Left            =   600
      TabIndex        =   33
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Overhead"
      Height          =   255
      Index           =   14
      Left            =   3480
      TabIndex        =   32
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material"
      Height          =   255
      Index           =   13
      Left            =   3480
      TabIndex        =   31
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Labor:"
      Height          =   255
      Index           =   12
      Left            =   600
      TabIndex        =   30
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Repeat Date, Comment and Account"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   29
      Top             =   1060
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Over/Short Account"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   28
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Cost"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   27
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   26
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   25
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter For Creating Lots Only:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   24
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Cost"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   23
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Lot Number"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   22
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment Quantity"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   20
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   19
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Top             =   720
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "MatlMMf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
Dim rdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim bOnLoad As Byte
Dim bGoodPart As Byte
Dim bView As Byte

Dim iTrans As Integer

'Dim sLots(100, 2) As String      ' 0 = lot#, 1 = user lot id (not used)
Dim sLots() As String      ' lot#
Dim sCreditAcct As String
Dim sDebitAcct As String
Private LotsExpire As Boolean

Private defaultLocation As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbAct_Click()
   FindAccount Me
   
End Sub

Private Sub cmbLot_Click()
   On Error Resume Next
   If cmbLot.ListIndex = -1 Then cmbLot.ListIndex = 0
   'lblNumber = sLots(cmbLot.ListIndex, 0)
   lblNumber = sLots(cmbLot.ListIndex)
   SelectThisLotQty
   
End Sub


Private Sub cmbLot_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   On Error Resume Next
   cmbLot = CheckLen(cmbLot, 40)
   For iList = 0 To cmbLot.ListCount - 1
      If cmbLot.List(iList) = cmbLot Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbLot = cmbLot.List(0)
   End If
   'lblNumber = sLots(cmbLot.ListIndex, 0)
   lblNumber = sLots(cmbLot.ListIndex)
   SelectThisLotQty
   
End Sub


Private Sub cmbPrt_Click()
   bGoodPart = GetPart()
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bView = 1 Then Exit Sub
   If Len(cmbPrt) > 0 Then bGoodPart = GetPart()
   
End Sub


Private Sub cmdAdd_Click()
   Dim cActivityQoh As Currency
   'On Error Resume Next
   If bGoodPart Then
      'No quantity
      If Val(txtQty) = 0 Then
         MsgBox "Invalid Quantity.", vbInformation, Caption
         txtQty.SetFocus
         Exit Sub
      End If
      
      'Force a Comment
      If Len(Trim(txtCmt)) = 0 Then
         MsgBox "Requires A Comment.", vbInformation, Caption
         txtCmt.SetFocus
         Exit Sub
      End If
      'Force them to correct negatives
      If Val(lblQoh) < 0 And Val(txtQty) < 0 Then
         MsgBox "There Is A Negative Quantity On Hand. " & vbCr _
            & "You Must Correct That Condition First.", vbInformation, Caption
         Exit Sub
      End If
      
      '1/13/04
      sDebitAcct = GetDebitAccount()
      sCreditAcct = GetCreditAccount()
      If optLot.Value = vbUnchecked Then
         cActivityQoh = GetActivityQuantity(Compress(cmbPrt))
         If cActivityQoh <> Val(lblQoh) Then RepairInventory Compress(cmbPrt), Val(lblQoh) - cActivityQoh
      End If
      If Val(txtQty) < 0 Then
         If optLot.Value = vbChecked Then
            'For Lots
            If Val(txtQty) < 0 And Val(txtLqoh) = 0 Then
               AdjustSubtractExistingLot  'OK
               Exit Sub
            End If
            If Abs(Val(txtQty)) > lblLotQty Then
               MsgBox "The Quantity To Be Subtracted Cannot Be" & vbCr _
                  & "Greater Than The Lot Quantity Remaining.", _
                  vbInformation, Caption
               Exit Sub
            Else
               AdjustSubtractExistingLot
               Exit Sub
            End If
         Else
            'For others
            If Abs(Val(txtQty)) > lblQoh Then
               MsgBox "The Quantity To Be Subtracted Cannot Be" & vbCr _
                  & "Greater Than The Lot Quantity On Hand.", _
                  vbInformation, Caption
               Exit Sub
            Else
               AdjustSubtractExistingLot
               Exit Sub
            End If
         End If
      End If
   End If
   
   'Adjust others
   AdjustAddExistingLot
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbPrt = ""
   
End Sub


Private Sub cmdFnd_Click()
   ViewParts.lblControl = "CMBPRT"
   ViewParts.txtPrt = cmbPrt 'txtPrt
   ViewParts.Show
   bView = 0
   
End Sub

Private Sub cmdFnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bView = 1
   
End Sub


Private Sub cmdHlp_Click()
   Dim l&
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5450"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub



Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yyyy"))
      If Left(sJournalID, 4) = "None" Then
         sJournalID = ""
         b = 1
      Else
         If sJournalID = "" Then b = 0 Else b = 1
      End If
      If b = 0 Then
         MouseCursor 0
         MsgBox "There Is No Open Inventory Journal For This Period.", _
            vbExclamation, Caption
         Sleep 500
         Unload Me
         Exit Sub
      End If
      FillAccounts
      If cmbAct.ListCount > 0 Then FindAccount Me
      If cmbPrt.Visible Then FillParts
      'Removed the annoying code 5/21/03
      ' If Len(Cur.CurrentPart) > 0 Then
      '     cmbPrt = Cur.CurrentPart
      '     bGoodPart = GetPart()
      ' Else
      If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
      ' End If
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PASTDCOST," _
          & "PAAVGCOST,PAUNITS,PAQOH,PAPRODCODE,PALOTQTYREMAINING," _
          & "PALOTTRACK,PATOOL, PALOTSEXPIRE, PALOCATION , CASE WHEN LEN(LOTNUMBER)=0 OR LOTNUMBER IS NULL THEN PASTDCOST Else LOTUNITCOST END AS UNITCOST " _
          & "From PartTable LEFT OUTER JOIN LohdTable ON LOTPARTREF = PARTREF " _
          & "WHERE (PARTREF= ? AND PATOOL=0)"
'   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PASTDCOST," _
'          & "PAAVGCOST,PAUNITS,PAQOH,PAPRODCODE,PALOTQTYREMAINING," _
'          & "PALOTTRACK,PATOOL, PALOTSEXPIRE, PALOCATION FROM PartTable WHERE (PARTREF= ? AND PATOOL=0)"
   Set rdoQry = New ADODB.Command
   rdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.Size = 30
   
   rdoQry.Parameters.Append AdoParameter1
   
   'RdoQry.MaxRows = 1
   cmbAct = GetSetting("Esi2000", "EsiAdmn", "AdjustAcct", cmbAct)
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If lblActDsc.ForeColor <> ES_RED Then
      SaveSetting "Esi2000", "EsiAdmn", "AdjustAcct", Trim(cmbAct)
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   On Error Resume Next
   Set AdoParameter1 = Nothing
   
   Set rdoQry = Nothing
   Set MatlMMf01a = Nothing
   
End Sub




Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtAvg.BackColor = Es_FormBackColor
   b = CheckLotStatus()
   If b = 1 Then
      z1(20).Visible = True
      txtLqoh.Visible = True
   End If
   txtStd.BackColor = Es_TextDisabled
   If ES_PARTCOUNT > 5000 Then
     ' txtPrt.Top = cmbPrt.Top
      cmdFnd.Top = cmbPrt.Top
      'cmbPrt.Visible = False
     ' txtPrt.Visible = True
      cmbPrt.TabIndex = 0
     ' cmdFnd.Visible = True
   End If
   
End Sub

Private Sub FillParts()
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PATOOL From PartTable Where " _
          & "(PAPRODCODE<>'BID' AND PATOOL=0 AND PAINACTIVE = 0 AND PAOBSOLETE = 0) ORDER BY PARTREF"
   LoadComboBox cmbPrt
   Exit Sub
   
DiaErr1:
   sProcName = "fillparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetPart() As Byte
   Dim RdoGet As ADODB.Recordset
   
   Dim sPartNumber As String
   Dim cLotQty As Currency
   sPartNumber = Compress(cmbPrt)
   On Error GoTo DiaErr1
   cmbLot.Clear
   rdoQry.Parameters(0).Value = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoGet, rdoQry, ES_KEYSET, True, 1)
   If bSqlRows Then
      With RdoGet
         GetPart = 1
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblLvl = "" & Format(!PALEVEL, "0")
         lblUom = "" & Trim(!PAUNITS)
         
         txtCst = "" & Format(!unitCost, ES_QuantityDataFormat) 'BBS Changed
         txtStd = "" & Format(!PASTDCOST, ES_QuantityDataFormat)
         txtAvg = "" & Format(!PAAVGCOST, ES_QuantityDataFormat)
         lblQoh = "" & Format(!PAQOH, ES_QuantityDataFormat)
         lblCode = "" & Trim(!PAPRODCODE)
         txtLqoh = "" & Format(!PALOTQTYREMAINING, ES_QuantityDataFormat)
         optLot = !PALOTTRACK
         If Not bOnLoad Then CheckBoxes True Else CheckBoxes False
         LotsExpire = IIf(!PALOTSEXPIRE = 0, False, True)
         lblExpDate.Visible = LotsExpire
         lblExpDateLabel.Visible = LotsExpire
         defaultLocation = !PALOCATION

      End With
      If optLot.Value = vbChecked Then
         GetPartLots
      Else
         lblNumber = ""
         lblDate = ""
         cmbLot.Text = "Disabled (Not Lot Tracked)"
         cmbLot.Enabled = False
         lblLotQty = "0.000"
      End If
   Else
      GetPart = 0
      CheckBoxes False
      txtStd = "0.000"
      txtCst = "0.000"
      lblQoh = "0.000"
      lblDsc = ""
      lblLvl = ""
      lblUom = ""
      MsgBox Trim(cmbPrt) & " Wasn't Found Or Doesn't Qualify.", _
                  vbInformation, Caption
   End If
   Set RdoGet = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub CheckBoxes(bOpen As Byte)
   If optRep.Value = vbUnchecked Then
      txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
      txtQty = "0.000"
      txtCmt = " "
   End If
   If bOpen Then
      txtQty.Enabled = True
      txtCst.Enabled = True
      ' txtLot.Enabled = True
      txtCmt.Enabled = True
      txtDte.Enabled = True
      cmbAct.Enabled = True
   Else
      txtQty.Enabled = False
      txtCst.Enabled = False
      'txtlot.Enabled = False
      txtCmt.Enabled = False
      txtDte.Enabled = False
      txtStd.Enabled = False
      cmbAct.Enabled = False
   End If
End Sub

Private Sub lblActdsc_Change()
   If Left(lblActDsc, 6) = "*** Ac" Then
      If sJournalID <> "" Then lblActDsc.ForeColor = ES_RED
   Else
      lblActDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optRep_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub cmbAct_LostFocus()
   cmbAct = CheckLen(cmbAct, 12)
   FindAccount Me
   
End Sub


'INREF2

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 40)
   txtCmt = StrCase(txtCmt)
   If Val(txtQty) <> 0 Then
      If Len(txtCmt) Then
         cmdAdd.Enabled = True
      Else
         cmdAdd.Enabled = False
      End If
   End If
   
End Sub


Private Sub txtCst_LostFocus()
   txtCst = CheckLen(txtCst, 9)
   txtCst = Format(Abs(Val(txtCst)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDateEx(txtDte)
   VerifyDate
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 10)
   txtQty = Format(Val(txtQty), ES_QuantityDataFormat)
   If Val(txtQty) <> 0 Then
      If Len(Trim(txtCmt)) Then
         cmdAdd.Enabled = True
      Else
         cmdAdd.Enabled = False
      End If
   End If
   
End Sub


Private Sub txtStd_LostFocus()
   txtStd = CheckLen(txtStd, 9)
   txtStd = Format(Abs(Val(txtStd)), ES_QuantityDataFormat)
   
End Sub



Private Sub VerifyDate()
   Dim lDte1 As Long
   Dim lDte2 As Long
   
   If Len(txtDte) Then
      lDte1 = DateValue(txtDte)
      lDte2 = DateValue(Format(ES_SYSDATE, "mm/dd/yyyy"))
      If lDte1 > lDte2 Then
         Beep
         txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
      End If
   End If
   
End Sub

'Lots 3/13/02 - only lots where balance is greater than zero
'Removed 3/15/03

Private Sub FillAccounts()
   Dim RdoAcct As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "Qry_FillLowAccounts"
   LoadComboBox cmbAct
   sSql = "SELECT COADJACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAcct, ES_FORWARD)
   If bSqlRows Then cmbAct = "" & Trim(RdoAcct!COADJACCT)
   If cmbAct.ListCount > 0 Then
      FindAccount Me
      lblActDsc = lblDsc
      lblDsc = ""
   End If
   Set RdoAcct = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillaccou"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetCreditAccount() As String
   Dim rdoAct As ADODB.Recordset
   
   Dim bType As Byte
   Dim sPcode As String
   
   On Error Resume Next
   sPcode = Compress(lblCode)
   bType = Val(lblLvl)
   'Part First
   sSql = "SELECT PAINVMATACCT FROM PartTable WHERE " _
          & "PARTREF='" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         GetCreditAccount = "" & Trim(.Fields(0))
         ClearResultSet rdoAct
      End With
   End If
   If GetCreditAccount = "" Then
      sSql = "SELECT PCINVMATACCT FROM PcodTable WHERE " _
             & "PCREF='" & Compress(sPcode) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            GetCreditAccount = "" & Trim(.Fields(0))
            ClearResultSet rdoAct
         End With
      End If
   End If
   sSql = "SELECT COINVMATACCT" & Trim(str(bType)) & " " _
          & "FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         If GetCreditAccount = "" Then GetCreditAccount = "" & Trim(.Fields(0))
         ClearResultSet rdoAct
      End With
   End If
   Set rdoAct = Nothing
   Exit Function
   
DiaErr1:
   'Just bail for now. May not have anything set
   'CurrError.Number = Err
   'CurrError.Description = Err.Description
   'DoModuleErrors Me
   On Error GoTo 0
   
End Function


Private Sub GetPartLots()
   Dim RdoLot As ADODB.Recordset
   
   Dim iRow As Integer
   Dim cLotAdj As Currency
   cmbLot.Enabled = False
   cmbLot.Clear
   Erase sLots
   ReDim sLots(1000)
   iRow = -1
   If optLot.Value = vbChecked Then cLotAdj = GetRemainingLotQty(Compress(cmbPrt), True)
   sSql = "SELECT LOTNUMBER,LOTUSERLOTID FROM LohdTable WHERE (LOTPARTREF='" _
          & Compress(cmbPrt) & "' AND LOTAVAILABLE=1) ORDER BY LOTUSERLOTID"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         Do Until .EOF
            iRow = iRow + 1
            AddComboStr cmbLot.hWnd, "" & Trim(!LOTUSERLOTID)
            
            'if lot array full, add another 1000 elements
            If iRow > UBound(sLots, 1) Then
               ReDim Preserve sLots(iRow + 999) As String
            End If
            
'            sLots(iRow, 0) = "" & Trim(!LotNumber)
'            sLots(iRow, 1) = "" & Trim(!LOTUSERLOTID)
            sLots(iRow) = "" & Trim(!lotNumber)
            .MoveNext
         Loop
         ClearResultSet RdoLot
      End With
   End If
   If cmbLot.ListCount > 0 Then
      cmbLot.Enabled = True
      cmbLot = cmbLot.List(0)
      'lblNumber = sLots(0, 0)
      lblNumber = sLots(0)
      SelectThisLotQty
   Else
      cmbLot = "No Lots Have Been Recorded"
      lblLotQty = "0.000"
   End If
   Set RdoLot = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpartlots"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'2/27/03 Negative Inventory
'Not in use

'Private Sub AdjustNegative(cOldQoh As Currency)
'End Sub

Private Sub SelectThisLotQty()
   Dim RdoSel As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTADATE,LOTREMAININGQTY," _
          & "LOTLOCATION, LOTEXPIRESON, LOTUNITCOST FROM LohdTable WHERE LOTNUMBER='" & Me.lblNumber & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoSel, ES_FORWARD)
   If bSqlRows Then
      With RdoSel
         If Not IsNull(!LOTREMAININGQTY) Then
            lblLotQty = Format(!LOTREMAININGQTY, ES_QuantityDataFormat)
            cmbLot.ToolTipText = "System Lot " & Trim(!lotNumber)
            lblNumber = "" & Trim(!lotNumber)
            lblDate = Format(!LotADate, "mm/dd/yyyy")
            LBLLotLoc = "" & Trim(!LOTLOCATION)
            txtCst = "" & Format(!LotUnitCost, ES_QuantityDataFormat) 'BBS Added for Ticket #59780
         Else
            lblLotQty = "0.000"
            lblNumber = ""
            cmbLot.ToolTipText = "Not A Valid Lot"
            lblDate = ""
            LBLLotLoc = ""
         End If
         
         If IsNull(!LOTEXPIRESON) Then
            Me.lblExpDate = ""
         Else
            Me.lblExpDate = Format(!LOTEXPIRESON, "mm/dd/yy")
         End If
         
         ClearResultSet RdoSel
      End With
   End If
   Set RdoSel = Nothing
   Exit Sub
   
DiaErr1:
   On Error GoTo 0
   lblLotQty = "0.000"
   cmbLot.ToolTipText = "Not A Valid Lot"
   lblNumber = ""
   
End Sub

'3/4/03 Adjust for reducing nots and no lots
'1/22/05 checked

Private Sub AdjustSubtractExistingLot()
   Dim bResponse As Byte
   
   Dim lCOUNTER As Long
   Dim lLOTRECORD As Long
   
   Dim cAdjQty As Currency
   Dim cCost As Currency
   Dim cOldQoh As Currency
   Dim cOldLotQty As Currency
   Dim cremAdjQty As Currency
   Dim cCurAdjQty As Currency
   
   Dim sLotNumber As String
   Dim sPartNumber As String
   Dim sMsg As String
   
   Dim vAdate As Variant
   
   sMsg = "Adjust Inventory For: " & cmbPrt & " " & vbCrLf _
          & "By Subtracting " & Abs(txtQty) & " From Inventory?" & vbCr _
          & "Lot Tracked Parts Will Use The Selected lot."
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
      Exit Sub
   End If
   
   vAdate = Format(ES_SYSDATE, "mm/dd/yyyy hh:mm")
   cAdjQty = Val(txtQty)
   sPartNumber = Compress(cmbPrt)
   cOldQoh = Val(lblQoh)
   cCost = Val(txtCst)
   sLotNumber = lblNumber
   
   'On Error Resume Next
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   lCOUNTER = GetLastActivity() + 1
   sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
          & "INADATE,INPDATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," _
          & "INNUMBER,INUSER,INLOTNUMBER) " _
          & "VALUES(19,'" & sPartNumber & "','Manual Adjustment','" & txtCmt & "'," _
          & "'" & Format(txtDte, "mm/dd/yyyy") & "','" & vAdate & "'," & cAdjQty _
          & "," & cAdjQty & "," & cCost & ",'" & sCreditAcct _
          & "','" & sDebitAcct & "'," & lCOUNTER & ",'" & sInitials _
          & "','" & sLotNumber & "')"
   clsADOCon.ExecuteSQL sSql
   'lots
   cOldLotQty = 0
   If optLot.Value = vbChecked Then
      'Lots checked
      
      If (cOldQoh < Abs(cAdjQty)) Then
         sSql = "UPDATE PartTable SET PAQOH=0,PALOTQTYREMAINING=0 " _
                & " WHERE PARTREF='" & sPartNumber & "'"
      Else
         sSql = "UPDATE PartTable SET PAQOH=PAQOH-" & Abs(cAdjQty) & "," _
                & "PALOTQTYREMAINING=PALOTQTYREMAINING-" & Abs(cAdjQty) & " " _
                & "WHERE PARTREF='" & sPartNumber & "'"
      End If
      
      clsADOCon.ExecuteSQL sSql
      
      lLOTRECORD = GetNextLotRecord(sLotNumber)
'      sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
'             & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
'             & "LOIACTIVITY,LOICOMMENT) " _
'             & "VALUES('" _
'             & sLotNumber & "'," & lLOTRECORD & ",19,'" & sPartNumber _
'             & "','" & vAdate & "'," & Trim(str(cAdjQty)) _
'             & "," & lCOUNTER & ",'" _
'             & "Manual Inventory Adjustment" & "')"
'      clsADOCon.ExecuteSQL sSql
 
      sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
             & "LOITYPE,LOIPARTREF,LOIADATE,LOIPDATE,LOIQUANTITY," _
             & "LOIACTIVITY,LOICOMMENT) " _
             & "VALUES('" _
             & sLotNumber & "'," & lLOTRECORD & ",19,'" & sPartNumber _
             & "','" & Format(txtDte, "mm/dd/yyyy") & "','" & vAdate & "'," & Trim(str(cAdjQty)) _
             & "," & lCOUNTER & ",'" _
             & "Manual Inventory Adjustment" & "')"
      clsADOCon.ExecuteSQL sSql
          
      sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY" _
             & cAdjQty & " WHERE LOTNUMBER='" & sLotNumber & "' "
      clsADOCon.ExecuteSQL sSql
   Else
      If Val(txtLqoh) >= Abs(cAdjQty) Then
         'see if there is one
         
         
         cremAdjQty = Abs(cAdjQty)
         While (cremAdjQty > 0)
            
            cCurAdjQty = SubTractFromOldLot(Abs(cremAdjQty), sLotNumber)
            
            If (cCurAdjQty > cremAdjQty) Then cCurAdjQty = cremAdjQty
            cremAdjQty = cremAdjQty - Abs(cCurAdjQty)
            
            If Abs(cCurAdjQty) > 0 Then
               ' make curAdjqty was negative
               If (cAdjQty < 0) Then cCurAdjQty = cCurAdjQty * -1
                  
               sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY" _
                      & cCurAdjQty & " WHERE LOTNUMBER='" & sLotNumber & "' "
               clsADOCon.ExecuteSQL sSql
               
               lLOTRECORD = GetNextLotRecord(sLotNumber)
               sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                      & "LOITYPE,LOIPARTREF,LOIADATE,LOIPDATE,LOIQUANTITY," _
                      & "LOIACTIVITY,LOICOMMENT) " _
                      & "VALUES('" _
                      & sLotNumber & "'," & lLOTRECORD & ",19,'" & sPartNumber _
                      & "','" & Format(txtDte, "mm/dd/yyyy") & "','" & vAdate & "'," & cCurAdjQty _
                      & "," & lCOUNTER & ",'" _
                      & "Manual Inventory Adjustment" & "')"
               clsADOCon.ExecuteSQL sSql
            End If
            
         Wend
      Else
         sLotNumber = GetNextLotNumber()
         'new lot and reset it
'         sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
'                & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
'                & "LOTUNITCOST,LOTDATECOSTED,LOTCOMMENTS) " _
'                & "VALUES('" _
'                & sLotNumber & "','Manual Adjustment-" & sLotNumber & "','" & sPartNumber _
'                & "','" & vAdate & "'," & Abs(cAdjQty) & ",0" _
'                & "," & cCost & ",'" & vAdate & "','" & txtCmt & "')"
         
         sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
                & "LOTPARTREF,LOTADATE, LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
                & "LOTUNITCOST,LOTDATECOSTED,LOTCOMMENTS) " _
                & "VALUES('" _
                & sLotNumber & "','Manual Adjustment-" & sLotNumber & "','" & sPartNumber _
                & "','" & Format(txtDte, "mm/dd/yy") & "','" & vAdate & "'," & Abs(cAdjQty) & ",0" _
                & "," & cCost & ",'" & Format(txtDte, "mm/dd/yy") & "','" & txtCmt & "')"
                
                
         clsADOCon.ExecuteSQL sSql
         
         sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                & "LOITYPE,LOIPARTREF,LOIADATE,LOIPDATE,LOIQUANTITY," _
                & "LOIACTIVITY,LOICOMMENT) " _
                & "VALUES('" _
                & sLotNumber & "',1,19,'" & sPartNumber _
                & "','" & Format(txtDte, "mm/dd/yyyy") & "','" & vAdate & "'," & Abs(cAdjQty) _
                & "," & lCOUNTER & ",'" _
                & "Manual Inventory Adjustment" & "')"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         
         sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                & "LOITYPE,LOIPARTREF,LOIADATE,LOIPDATE,LOIQUANTITY," _
                & "LOIACTIVITY,LOICOMMENT) " _
                & "VALUES('" _
                & sLotNumber & "',2,19,'" & sPartNumber _
                & "','" & Format(txtDte, "mm/dd/yyyy") & "','" & vAdate & "'," & cAdjQty _
                & "," & lCOUNTER & ",'" _
                & "Manual Inventory Adjustment" & "')"
         clsADOCon.ExecuteSQL sSql ' rdExecDirect

      End If
      
      If (cOldQoh < Abs(cAdjQty)) Then
         sSql = "UPDATE PartTable SET PAQOH=0, PALOTQTYREMAINING=0 " _
                & " WHERE PARTREF='" & sPartNumber & "'"
      Else
         sSql = "UPDATE PartTable SET PAQOH=PAQOH-" & Abs(cAdjQty) & "," _
                & "PALOTQTYREMAINING=PALOTQTYREMAINING-" & Abs(cAdjQty) & " " _
                & "WHERE PARTREF='" & sPartNumber & "'"
      End If
      
      clsADOCon.ExecuteSQL sSql
      UpdateWipColumns lCOUNTER
   End If
   If clsADOCon.ADOErrNum = 0 Then
      SysMsg "Transaction Complete", True
      clsADOCon.CommitTrans
      bGoodPart = GetPart()
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox "Couldn't complete The Transaction.", _
         vbExclamation, Caption
   End If
   
End Sub

'1/22/05 Revised Debit and Credits (swapped)

Private Sub AdjustAddExistingLot()
   Dim bResponse As Byte
   Dim bNewLot As Byte
   
   Dim lCOUNTER As Long
   Dim lLOTRECORD As Long
   
   Dim cAdjQty As Currency
   Dim cCost As Currency
   
   Dim sLotNumber As String
   Dim sPartNumber As String
   Dim sMsg As String
   
   Dim vAdate As Variant
   
   If optLot.Value = vbChecked Then
      sMsg = "This Part Requires Lot Tracking. Do You Wish To " & vbCrLf _
             & "Use The Lot Number Selected?" & vbCr _
             & "Otherwise A New Lot Will Be Created."
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbNo Then
         bNewLot = 1
      Else
         bNewLot = 0
      End If
   Else
      bNewLot = 1
   End If
   sMsg = "Adjust Inventory For: " & cmbPrt & " " & vbCrLf _
          & "By Adding " & Abs(txtQty) & " To Inventory?" & vbCr
   If bNewLot = 1 Then
      sMsg = sMsg & "A New Lot Number Will Be Created."
   Else
      sMsg = sMsg & "Lot Tracked Parts Will Use Sys: " & lblNumber & "."
   End If
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
      Exit Sub
   End If
   
   If bNewLot = 1 Then
      'Create a new lot
      sLotNumber = GetNextLotNumber()
   Else
      'Use a current lot to adjust
      sLotNumber = lblNumber
   End If
   
   If sJournalID <> "" Then iTrans = GetNextTransaction(sJournalID)
   vAdate = Format(ES_SYSDATE, "mm/dd/yyyy hh:mm")
   cAdjQty = Val(txtQty)
   sPartNumber = Compress(cmbPrt)
   cCost = Val(txtCst)
   
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   lCOUNTER = GetLastActivity() + 1
   sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
          & "INADATE,INPDATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," _
          & " INNUMBER,INUSER, INLOTNUMBER) " _
          & "VALUES(19,'" & sPartNumber & "','Manual Adjustment','" & txtCmt & "'," _
          & "'" & Format(txtDte, "mm/dd/yyyy") & "','" & vAdate & "'," & cAdjQty _
          & "," & cAdjQty & "," & cCost & ",'" & sCreditAcct _
          & "','" & sDebitAcct & "'," & lCOUNTER & ",'" & sInitials _
          & "','" & sLotNumber & "')"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "UPDATE PartTable SET PAQOH=PAQOH+" & Abs(cAdjQty) & "," _
          & "PALOTQTYREMAINING=PALOTQTYREMAINING+" & Abs(cAdjQty) & " " _
          & "WHERE PARTREF='" & sPartNumber & "'"
   clsADOCon.ExecuteSQL sSql
   
   If bNewLot = 1 Then
      lLOTRECORD = 1
'      sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
'             & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
'             & "LOTUNITCOST,LOTDATECOSTED,LOTCOMMENTS) " _
'             & "VALUES('" _
'             & sLotNumber & "','Manual Adjustment-" & sLotNumber & "','" & sPartNumber _
'             & "','" & vAdate & "'," & cAdjQty & "," & cAdjQty & "" _
'             & "," & cCost & ",'" & vAdate & "','" & txtCmt & "')"
   
      sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
             & "LOTPARTREF,LOTADATE, LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
             & "LOTUNITCOST,LOTDATECOSTED,LOTCOMMENTS) " _
             & "VALUES('" _
             & sLotNumber & "','Manual Adjustment-" & sLotNumber & "','" & sPartNumber _
             & "','" & Format(txtDte, "mm/dd/yy") & "','" & vAdate & "'," & cAdjQty & "," & cAdjQty & "" _
             & "," & cCost & ",'" & Format(txtDte, "mm/dd/yy") & "','" & txtCmt & "')"
   
   Else
      lLOTRECORD = GetNextLotRecord(sLotNumber)
'Rev 66 ticket 9565 - Don't update original qty
'      sSql = "UPDATE LohdTable SET LOTORIGINALQTY=LOTORIGINALQTY+" _
'             & cAdjQty & ",LOTREMAININGQTY=LOTREMAININGQTY+" _
'             & cAdjQty & " WHERE LOTNUMBER='" & sLotNumber & "'"
      sSql = "UPDATE LohdTable" & vbCrLf _
         & "SET LOTREMAININGQTY = LOTREMAININGQTY  + " & cAdjQty & vbCrLf _
         & "WHERE LOTNUMBER = '" & sLotNumber & "'"
   End If
   clsADOCon.ExecuteSQL sSql
   
'   sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
'          & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
'          & "LOIACTIVITY,LOICOMMENT) " _
'          & "VALUES('" _
'          & sLotNumber & "'," & lLOTRECORD & ",19,'" & sPartNumber _
'          & "','" & vAdate & "'," & cAdjQty _
'          & "," & lCOUNTER & ",'" _
'          & "Manual Inventory Adjustment" & "')"
   
   sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
          & "LOITYPE,LOIPARTREF,LOIADATE, LOIPDATE,LOIQUANTITY," _
          & "LOIACTIVITY,LOICOMMENT) " _
          & "VALUES('" _
          & sLotNumber & "'," & lLOTRECORD & ",19,'" & sPartNumber _
          & "','" & Format(txtDte, "mm/dd/yyyy") & "','" & vAdate & "'," & cAdjQty _
          & "," & lCOUNTER & ",'" _
          & "Manual Inventory Adjustment" & "')"
                  
   clsADOCon.ExecuteSQL sSql
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      If bNewLot = 1 And optLot.Value = vbChecked Then
         MsgBox "The Adjustment Was Made And Lot Number " & vbCr _
            & "System: " & sLotNumber & " Was Created." & vbCr _
            & "You May Edit Some Features Now.", _
            vbInformation, Caption
         
         LotEdit.txtUnitCost = txtCst
         LotEdit.txtLong = txtCmt
         LotEdit.lblNumber = "Manual Adjustment-" & sLotNumber
         LotEdit.lblPart = cmbPrt
         LotEdit.lblDate = Format(ES_SYSDATE, "mm/dd/yyyy")
         LotEdit.lblTime = Format(ES_SYSDATE, "hh:mm")
         LotEdit.lblNumber = sLotNumber
         LotEdit.Show 1
      End If
      AverageCost sPartNumber
      UpdateWipColumns lCOUNTER
      SysMsg "Inventory Was Adjusted.", True
      bGoodPart = GetPart()
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      
      MsgBox "The Inventory Transaction Was Not Completed.", _
         vbExclamation, Caption
   End If
   
End Sub

'4/30/03 find an old lot and subtract
'We'll gamble on a lot

Private Function SubTractFromOldLot(cAdjust As Currency, sLotNum As String) As Currency
   Dim RdoOld As ADODB.Recordset
   Dim sPartNum As String

   sPartNum = Compress(cmbPrt)
   sSql = "SELECT LOTNUMBER,LOTPARTREF,LOTREMAININGQTY " _
          & "FROM LohdTable WHERE LOTPARTREF='" & sPartNum _
          & "' AND LOTREMAININGQTY > 0 AND LOTAVAILABLE = 1 ORDER BY LOTNUMBER ASC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOld, ES_FORWARD)
   If bSqlRows Then
      With RdoOld
         sLotNum = "" & Trim(!lotNumber)
         SubTractFromOldLot = !LOTREMAININGQTY
         ClearResultSet RdoOld
      End With
   Else
      SubTractFromOldLot = 0
   End If
End Function

'Private Function SubTractFromOldLot(cAdjust As Currency, sLotNum As String) As Currency
'   Dim RdoOld As ADODB.Recordset
'
'   Dim sPartNum As String
'
'   sPartNum = Compress(cmbPrt)
'   sSql = "SELECT LOTNUMBER,LOTPARTREF,LOTREMAININGQTY " _
'          & "FROM LohdTable WHERE (LOTPARTREF='" & sPartNum _
'          & "' AND LOTREMAININGQTY >=" & cAdjust & ")"
'   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOld, ES_FORWARD)
'   If bSqlRows Then
'      With RdoOld
'         sLotNum = "" & Trim(!lotNumber)
'         ClearResultSet RdoOld
'      End With
'      SubTractFromOldLot = cAdjust
'   Else
'      SubTractFromOldLot = 0
'   End If
'   Set RdoOld = Nothing
'End Function

Private Function GetDebitAccount() As String
   Dim rdoAct As ADODB.Recordset
   
   Dim bType As Byte
   Dim sPcode As String
   
   On Error Resume Next
   sPcode = Compress(lblCode)
   bType = Val(lblLvl)
   GetDebitAccount = ""
   
   'Default Over/Short
   sSql = "SELECT COADJACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         If Not IsNull(!COADJACCT) Then _
                       GetDebitAccount = "" & Trim(!COADJACCT)
         ClearResultSet rdoAct
      End With
   End If
   Set rdoAct = Nothing
   If GetDebitAccount <> "" Then Exit Function
   'Part First
   sSql = "SELECT PACGSMATACCT FROM PartTable WHERE " _
          & "PARTREF='" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         GetDebitAccount = "" & Trim(.Fields(0))
         ClearResultSet rdoAct
      End With
   End If
   Set rdoAct = Nothing
   If GetDebitAccount = "" Then
      sSql = "SELECT PCCGSMATACCT FROM PcodTable WHERE " _
             & "PCREF='" & Compress(sPcode) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            GetDebitAccount = "" & Trim(.Fields(0))
            ClearResultSet rdoAct
         End With
      End If
      Set rdoAct = Nothing
   End If
   sSql = "SELECT COCGSMATACCT" & Trim(str(bType)) & " " _
          & "FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         If GetDebitAccount = "" Then GetDebitAccount = "" & Trim(.Fields(0))
         ClearResultSet rdoAct
      End With
   End If
   Set rdoAct = Nothing
   Exit Function
   
DiaErr1:
   'Just bail for now. May not have anything set
   'CurrError.Number = Err
   'CurrError.Description = Err.Description
   'DoModuleErrors Me
   On Error GoTo 0
   
End Function
