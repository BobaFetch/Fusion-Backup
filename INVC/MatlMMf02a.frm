VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form MatlMMf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Inventory Activity Standard Costs"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "MatlMMf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbAct 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "8"
      ToolTipText     =   "Valid Inventory Activities For Update"
      Top             =   2040
      Width           =   2652
   End
   Begin VB.ComboBox txtBeg 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   4
      ToolTipText     =   "Update Selected Rows And Apply Changes"
      Top             =   2040
      Width           =   875
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   288
      Left            =   1560
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "The List Does Not Contain Lot Tracked Part Numbers"
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4920
      FormDesignWidth =   6870
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity Type (Select)"
      Height          =   252
      Index           =   12
      Left            =   240
      TabIndex        =   26
      Top             =   2040
      Width           =   1932
   End
   Begin VB.Line Line1 
      X1              =   5640
      X2              =   6720
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label lblHours 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   5640
      TabIndex        =   25
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Hours"
      ForeColor       =   &H80000008&
      Height          =   288
      Index           =   11
      Left            =   3600
      TabIndex        =   24
      Top             =   2520
      Width           =   1272
   End
   Begin VB.Label lblOvrhd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   5640
      TabIndex        =   23
      Top             =   3960
      Width           =   1092
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Overhead Cost"
      ForeColor       =   &H80000008&
      Height          =   288
      Index           =   10
      Left            =   3600
      TabIndex        =   22
      Top             =   3960
      Width           =   1272
   End
   Begin VB.Label lblLabor 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   5640
      TabIndex        =   21
      Top             =   2880
      Width           =   1092
   End
   Begin VB.Label lblMatl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   5640
      TabIndex        =   20
      Top             =   3600
      Width           =   1092
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Cost"
      ForeColor       =   &H80000008&
      Height          =   288
      Index           =   9
      Left            =   3600
      TabIndex        =   19
      Top             =   2880
      Width           =   1272
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Material Cost"
      ForeColor       =   &H80000008&
      Height          =   288
      Index           =   8
      Left            =   3600
      TabIndex        =   18
      Top             =   3600
      Width           =   1272
   End
   Begin VB.Label lblExp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   5640
      TabIndex        =   17
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Cost"
      ForeColor       =   &H80000008&
      Height          =   288
      Index           =   7
      Left            =   3600
      TabIndex        =   16
      Top             =   3240
      Width           =   1272
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   " From"
      Height          =   252
      Index           =   6
      Left            =   1800
      TabIndex        =   15
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   252
      Index           =   5
      Left            =   3600
      TabIndex        =   14
      Top             =   1680
      Width           =   732
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Activity Rows (Actual):"
      ForeColor       =   &H80000008&
      Height          =   288
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   2712
   End
   Begin VB.Label lblStd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   5640
      TabIndex        =   12
      Top             =   4440
      Width           =   1092
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Cost"
      ForeColor       =   &H80000008&
      Height          =   288
      Index           =   3
      Left            =   3600
      TabIndex        =   11
      Top             =   4440
      Width           =   1272
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   5640
      TabIndex        =   10
      Top             =   840
      Width           =   492
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   288
      Index           =   1
      Left            =   4920
      TabIndex        =   9
      Top             =   840
      Width           =   1272
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   288
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   1272
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1272
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1560
      TabIndex        =   6
      Top             =   1200
      Width           =   3012
   End
End
Attribute VB_Name = "MatlMMf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'9/23/04 new
'9/20/05 Overhauled (Added Standards, Inva Types)
'See UpdateStandardCosts2005-09-20.txt (email to Larry/Terry).
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   GetPartInfo
   
End Sub


Private Sub cmbPrt_LostFocus()
   GetPartInfo
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1650
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   If Val(Left(cmbAct, 2)) = 0 Then
      MsgBox "Requires A Valid Inventory Activity From The List.", _
         vbInformation, Caption
      cmbAct = cmbAct.List(0)
   Else
      If Val(lblStd) = 0 Then
         bResponse = MsgBox("The Standard Cost Is Zero. Continue?", _
                     ES_NOQUESTION, Caption)
         If bResponse = vbNo Then
            CancelTrans
            Exit Sub
         End If
      End If
      UpdateActivity
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillTypes
      FillCombo
      bOnLoad = 0
   End If
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
   FormUnload
   Set MatlMMf02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 5)
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE " _
          & "PALOTTRACK=0 AND PAINACTIVE = 0 AND PAOBSOLETE = 0 ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 6) = "*** Pa" Then lblDsc.ForeColor = ES_RED _
           Else lblDsc.ForeColor = vbBlack
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   txtBeg = CheckDateEx(txtBeg)
   
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   txtBeg = CheckDateEx(txtBeg)
   
End Sub



Private Sub GetPartInfo()
   Dim RdoInf As ADODB.Recordset
   sSql = "SELECT * FROM PartTable WHERE PARTREF='" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInf, ES_FORWARD)
   If bSqlRows Then
      With RdoInf
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblType = "" & Trim(!PALEVEL)
         lblHours = Format(!PALEVHRS, ES_QuantityDataFormat)
         lblLabor = Format(!PALEVLABOR, ES_QuantityDataFormat)
         lblExp = Format(!PALEVEXP, ES_QuantityDataFormat)
         lblMatl = Format(!PALEVMATL, ES_QuantityDataFormat)
         lblOvrhd = Format(!PALEVOH, ES_QuantityDataFormat)
         lblStd = Format(!PASTDCOST, ES_QuantityDataFormat)
         
         cmdUpd.Enabled = True
         ClearResultSet RdoInf
      End With
   Else
      lblDsc = "*** Part Number Wasn't Found ***"
      cmdUpd.Enabled = False
   End If
   Set RdoInf = Nothing
   
End Sub

Private Sub UpdateActivity()
   Dim RdoStd As ADODB.Recordset
   Dim bResponse As Byte
   Dim bType As Byte
   Dim lRows As Long
   Dim cCost As Currency
   
   bType = Val(Left(cmbAct, 2))
   sSql = "SELECT INPART,INADATE FROM InvaTable WHERE " _
          & "(INADATE BETWEEN '" & txtBeg & " 00:00' AND '" _
          & txtEnd & " 23:59') AND INPART='" & Compress(cmbPrt) & "' " _
          & "AND INTYPE=" & Val(bType) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStd, ES_FORWARD)
   If bSqlRows Then
      With RdoStd
         Do Until .EOF
            lRows = lRows + 1
            .MoveNext
         Loop
         ClearResultSet RdoStd
      End With
   End If
   Set RdoStd = Nothing
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   bResponse = MsgBox(lRows & " Activity Row(S) Will Be Updated. Continue?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      cCost = GetPartCost(Compress(cmbPrt))
      sSql = "UPDATE InvaTable SET INAMT=" & cCost & "," _
             & "INTOTHRS=" & Val(lblHours) & "," _
             & "INTOTLABOR=" & Val(lblLabor) & "," _
             & "INTOTEXP=" & Val(lblExp) & "," _
             & "INTOTMATL=" & Val(lblMatl) & "," _
             & "INTOTOH=" & Val(lblOvrhd) & " " _
             & "WHERE (INADATE BETWEEN '" & txtBeg & " 00:00' AND '" _
             & txtEnd & " 23:59') AND INPART='" & Compress(cmbPrt) & "' " _
             & "AND INTYPE=" & Val(bType) & " "
      'MsgBox sSql
      clsADOCon.ExecuteSQL sSql
      AverageCost Compress(cmbPrt)
      If clsADOCon.ADOErrNum = 0 Then
         SysMsg "Activity Updated..", True
      Else
         MsgBox "Couldn't Successfully Update..", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Sub FillTypes()
   cmbAct.AddItem "19 - Manual Adjustment"
   cmbAct.AddItem "32 - Inventory Transfer"
   cmbAct.AddItem "10 - Actual Pick"
   cmbAct.AddItem "13 - Pick Surplus"
   cmbAct.AddItem "21 - Restocked Item"
   cmbAct.AddItem "22 - Scrapped Pick Item"
   cmbAct.AddItem "23 - Pick Substitute"
   cmbAct.AddItem "15 - PO Receipt"
   cmbAct.AddItem "16 - Canceled PO Receipt"
   cmbAct.AddItem "17 - Invoiced PO Item"
   cmbAct.AddItem "03 - Shipped Item (No Packing Slip)"
   cmbAct.AddItem "04 - Returned Item"
   cmbAct.AddItem "25 - Packing Slip (Out)"
   cmbAct.AddItem "33 - Canceled Packing Slip Item (In)"
   cmbAct.AddItem "06 - Completed MO"
   cmbAct.AddItem "07 - Closed MO"
   cmbAct.AddItem "38 - Canceled MO Completion"
   cmbAct = cmbAct.List(0)
   
   
End Sub
