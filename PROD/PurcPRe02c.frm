VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PurcPRe02c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Order Service Items"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbOrigDueDte 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3120
      TabIndex        =   40
      Tag             =   "4"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddStat 
      Height          =   375
      Left            =   6360
      Picture         =   "PurcPRe02c.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Status Code"
      Top             =   5400
      Width           =   375
   End
   Begin VB.ComboBox cmbAccount 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3720
      TabIndex        =   37
      ToolTipText     =   "Select Account From List"
      Top             =   6600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox optDockInsted 
      Caption         =   "Check1"
      Height          =   255
      Left            =   6480
      TabIndex        =   34
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRe02c.frx":048F
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optRcd 
      Alignment       =   1  'Right Justify
      Caption         =   "Received "
      Enabled         =   0   'False
      Height          =   255
      Left            =   4180
      TabIndex        =   30
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox optSel 
      Caption         =   "Select"
      Height          =   255
      Left            =   5880
      TabIndex        =   29
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Retrieve Service Parts"
      Top             =   2640
      Width           =   915
   End
   Begin VB.CommandButton cmdTrm 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Cancel The Current PO Item"
      Top             =   2040
      Width           =   915
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   315
      Left            =   6000
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Add A New PO Item (Select Service Item First)"
      Top             =   1680
      Width           =   915
   End
   Begin VB.TextBox txtCmt 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "9"
      ToolTipText     =   "Comments (2048 Chars Max)"
      Top             =   5400
      Width           =   4335
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "PurcPRe02c.frx":0C3D
      DownPicture     =   "PurcPRe02c.frx":15AF
      Enabled         =   0   'False
      Height          =   350
      Left            =   5880
      Picture         =   "PurcPRe02c.frx":1F21
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Standard Comments"
      Top             =   5400
      Width           =   350
   End
   Begin VB.TextBox txtLot 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Lot Quantity"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CheckBox optIns 
      Alignment       =   1  'Right Justify
      Caption         =   "Inspection Required?"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      ToolTipText     =   "Inspect This Item?"
      Top             =   5040
      Width           =   1995
   End
   Begin VB.ComboBox txtDue 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4800
      TabIndex        =   5
      Tag             =   "4"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.ComboBox cmbMon 
      Height          =   288
      Left            =   720
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Assign To MO"
      Top             =   2640
      Width           =   3195
   End
   Begin VB.ComboBox cmbRun 
      Height          =   288
      Left            =   5040
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Select Run Number"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Enter Quantity"
      Top             =   3876
      Width           =   1095
   End
   Begin VB.TextBox txtPrc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Price"
      Top             =   3876
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   6000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6945
      FormDesignWidth =   7005
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   2292
      Left            =   360
      TabIndex        =   31
      ToolTipText     =   "Click To Select Or Scroll And Press Enter"
      Top             =   120
      Width           =   5532
      _ExtentX        =   9763
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      HighLight       =   2
      ScrollBars      =   2
   End
   Begin VB.Label lblOrigShip 
      Caption         =   "Orig Due Date"
      Height          =   255
      Left            =   2040
      TabIndex        =   41
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblPua 
      Caption         =   "Purchase Account"
      Height          =   255
      Left            =   2160
      TabIndex        =   38
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDte 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblPrc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4800
      TabIndex        =   35
      ToolTipText     =   "Lot cost Price"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Dummy 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblPon 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   6360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Quantity If Not A Unit Price"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   24
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label lblOpno 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   5040
      TabIndex        =   23
      Top             =   3000
      Width           =   492
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Op No"
      Height          =   252
      Index           =   11
      Left            =   4200
      TabIndex        =   22
      Top             =   3000
      Width           =   852
   End
   Begin VB.Label lblMoDesc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   720
      TabIndex        =   21
      Top             =   3000
      Width           =   2892
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO"
      Height          =   492
      Index           =   7
      Left            =   240
      TabIndex        =   20
      Top             =   2640
      Width           =   1092
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   252
      Index           =   8
      Left            =   4200
      TabIndex        =   19
      Top             =   2640
      Width           =   852
   End
   Begin VB.Label lblServPart 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1440
      TabIndex        =   18
      Top             =   3876
      Width           =   2892
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity           "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   17
      Top             =   3636
      Width           =   1092
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price/Due                    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   4800
      TabIndex        =   16
      Top             =   3648
      Width           =   1572
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1440
      TabIndex        =   15
      Top             =   4200
      Width           =   2892
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6000
      TabIndex        =   14
      Top             =   3876
      Width           =   372
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   1440
      TabIndex        =   13
      Top             =   3636
      Width           =   3132
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   612
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1080
      TabIndex        =   11
      Top             =   3360
      Width           =   252
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Height          =   288
      Left            =   720
      TabIndex        =   10
      Top             =   3360
      Width           =   372
   End
End
Attribute VB_Name = "PurcPRe02c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of          ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'12/15/03 New (split off at request of LH)
'5/19/04 Allow Reprice of Received Items
'9/29/04 Reformatted values to prevent overflow
'10/7/04 ES_PurchasedDataFormat
'3/7/05 Added PIENTERED
'8/31/06 Expanded Grid
'1/31/07 Fixed cmbMon clearing 7.2.3
'3/28/07 Fixed loading first run in preloaded cmbRun 7.3.2
Option Explicit
Dim rdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim ADOParameter2 As ADODB.Parameter
Dim AdoParameter3 As ADODB.Parameter

Dim RdoEdit As ADODB.Recordset

Public bAllowPrice As Byte 'Allow repricing after receipt
Dim bOnLoad As Byte
Dim bCantCancel As Byte
Dim bDeleteThis As Byte
Dim bGoodItem As Byte
Dim bGoodMo As Byte
Dim bGoodRun As Byte
Dim bItemSel As Byte
Dim iRows As Integer
Dim sOldMoPart As String
Dim bFieldChanged As Byte


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbAccount_LostFocus()
   cmbAccount = CheckLen(cmbAccount, 12)
   With RdoEdit
      On Error Resume Next
      !PIACCOUNT = cmbAccount
      .Update
   End With
End Sub

Private Sub cmbMon_Click()
   If bOnLoad = 0 Then If sOldMoPart <> cmbMon Then GetRuns
   
End Sub


Private Sub cmbMon_LostFocus()
   cmbMon = CheckLen(cmbMon, 30)
   If cmbMon.Enabled Then
    If Len(cmbMon) > 0 And Not IsValidMONumber(cmbMon) Then
        MsgBox "Invalid MO Entered", vbExclamation
        cmbMon.SetFocus
        Exit Sub

    End If
   End If
   
    
   
   If Len(Trim(cmbMon)) Then
      If bOnLoad = 0 Then If sOldMoPart <> cmbMon Then GetRuns
      sOldMoPart = cmbMon
   Else
      lblMoDesc = "*** MO Part Is Blank ***"
      cmbRun.Clear
      sOldMoPart = ""
      bGoodMo = 0
   End If
   If bGoodMo = 1 Then
      With RdoEdit
         On Error Resume Next
         !PIRUNPART = Compress(cmbMon)
         .Update
      End With
   End If
   
End Sub


Private Sub cmbRun_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   '    cmbRun = Abs(Val(cmbRun))
   '        For iList = 0 To cmbRun.ListCount - 1
   '            If cmbRun = cmbRun.List(iList) Then bByte = 1
   '        Next
   '        If bByte = 1 Then
   '            cmdSel.Enabled = True
   '            bGoodRun = 1
   '        Else
   '            cmdSel.Enabled = False
   '            bGoodRun = 0
   '        End If
   '    If bGoodMo = 1 Then
   '        With RdoEdit
   '            On Error Resume Next
   '            .Edit
   '            !PIRUNPART = Compress(cmbMon)
   '            !PIRUNNO = Val(cmbRun)
   '            .Update
   '        End With
   '    End If
   
   
End Sub


Private Sub cmdAddStat_Click()
    StatusCode.lblSCTypeRef = "PO"
    StatusCode.txtSCTRef = lblPon
    StatusCode.LableRef1 = "Item"
    StatusCode.lblSCTRef1 = lblItm
    StatusCode.lblSCTRef2 = lblRev
    StatusCode.lblStatType = "POI"
    StatusCode.lblSysCommIndex = 1 ' The index in the Sys Comment "MO Comments"
    StatusCode.txtCurUser = cUR.CurrentUser
    
    StatusCode.Show
End Sub

Private Sub cmdCan_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If bItemSel = 1 Then
      bDeleteThis = 1
      If Trim(cmbMon) = "" Or Val(cmbRun) = 0 Or lblServPart = "" Then
         sMsg = "You Have Not Properly Select An Item." & vbCr _
                & "This Item Will Be Deleted From The PO. " & vbCr _
                & "Continue To Close Anyway?"
         bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
         If bResponse = vbYes Then Unload Me Else bDeleteThis = 0
      Else
         bDeleteThis = 0
         Unload Me
      End If
   Else
      bDeleteThis = 0
      Unload Me
   End If
   
End Sub



Private Sub cmdComments_Click()
   If cmdComments Then
      'See List For Index
      txtCmt.SetFocus
      SysComments.lblListIndex = 1
      SysComments.Show
      cmdComments = False
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4306
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdNew_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Add A Service Item For An MO Routing Item?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      cmbMon.Enabled = True
      cmbRun.Enabled = True
      AddPOItem
      ' MM 5/6/2010 The previous item was updating
      ' So moved the data set to current row.
      'GetThisItem
      ' Create new record set and start updating the new one.
      rdoQry.Parameters(0).Value = Val(lblPon)
      rdoQry.Parameters(1).Value = Trim(lblRev)
      rdoQry.Parameters(2).Value = Val(lblItm)
      
'      rdoQry(0) = Val(lblPon)
'      rdoQry(1) = Trim(lblRev)
'      rdoQry(2) = Val(lblItm)

'      GetQuerySet RdoEdit, rdoQry, ES_KEYSET
       clsADOCon.GetQuerySet RdoEdit, rdoQry, ES_KEYSET, True
   Else
      CancelTrans
   End If
   
End Sub

Private Sub cmdSel_Click()
   optSel.Value = vbChecked
   PurcPRe02d.lblMon = cmbMon
   PurcPRe02d.lblRun = cmbRun
   PurcPRe02d.Show
   
End Sub

Private Sub cmdTrm_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
    If optIns.Value = vbChecked Then
      If optDockInsted.Value = vbChecked Then
        MsgBox "You May Not Cancel A Dock Inspected Item.", _
            vbInformation, Caption
        Exit Sub
      End If
    End If
   
   
   bDeleteThis = 1
   sMsg = "You Have Chosen To Cancel Item " & lblItm & ". This Function" & vbCr _
          & "Cannot Be Reversed. Continue To Cancel Anyway?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sSql = "delete from PoitTable" & vbCrLf _
             & "where PINUMBER=" & Val(lblPon) _
             & " and PIITEM=" & Val(lblItm) _
             & " and PIREV='" & lblRev & "'"
      
      clsADOCon.ExecuteSQL sSql
      
      'reduce quantity purchased
      sSql = "UPDATE RnopTable" & vbCrLf _
             & "set OPPURCHASED=0," & vbCrLf _
             & "OPPONUMBER=0," & vbCrLf _
             & "OPPOITEM=''" & vbCrLf _
             & "WHERE OPREF='" & Compress(cmbMon) & "' AND " _
             & "OPRUN=" & Val(cmbRun) & " AND " _
             & "OPNO=" & Val(lblOpno) & vbCrLf _
             & "and not exists( select PINUMBER FROM PoitTable where PINUMBER=" & Val(lblPon) & vbCrLf _
             & "and PIITEM=" & Val(lblItm) & " )"
      
      clsADOCon.ExecuteSQL sSql
      
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      Grd.row = 2
      Grd.Col = 0
      Grd.Text = ""
      Grd.Col = 1
      Grd.Text = ""
      Grd.Col = 2
      Grd.Text = ""
      Grd.Col = 3
      Grd.Text = ""
      FillGrid
   Else
      CancelTrans
   End If
   
End Sub

Private Sub Form_Activate()
    Dim b As Byte
    
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      ES_PurchasedDataFormat = GetPODataFormat()
      bItemSel = 0
      FillCombo
      FillGrid
      ' Fill the ship date from the previous form.
      If iRows = 0 Then txtDue = lblDte
      
      b = CheckPoAccounts()
      FillAccounts
      If b = 1 Then
         lblPua.Visible = True
         cmbAccount.Visible = True
      End If
       
      
   End If
   optSel.Value = vbUnchecked
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   If optSel.Value = vbUnchecked Then Unload Me
   
End Sub


Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move PurcPRe02a.Left + 300, PurcPRe02a.Top + 1000
   FormatControls
   lblDte = PurcPRe02a.txtSdt
   If Trim(lblDte) = "" Then lblDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   sSql = "SELECT * FROM PoitTable WHERE (PINUMBER= ? " _
          & " AND PIREV= ? AND PIITEM= ? )"
   Set rdoQry = New ADODB.Command
   rdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adInteger
   rdoQry.Parameters.Append AdoParameter1
   
   Set ADOParameter2 = New ADODB.Parameter
   ADOParameter2.Type = adChar
   ADOParameter2.SIZE = 2
   rdoQry.Parameters.Append ADOParameter2
   
   Set AdoParameter3 = New ADODB.Parameter
   AdoParameter3.Type = adInteger
   rdoQry.Parameters.Append AdoParameter3
   
   
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .row = 0
      .Col = 0
      .Text = "Item"
      .ColWidth(0) = 650
      .Col = 1
      .Text = "Rev"
      .ColWidth(1) = 400
      .Col = 2
      .Text = "Service Part Number"
      .ColWidth(2) = 3450
      .Col = 3
      .Text = "Quantity"
      .ColWidth(3) = 950
      .Col = 0
   End With
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bDeleteThis = 1 Then
      On Error Resume Next
      sSql = "DELETE from PoitTable where (PINUMBER=" & Val(lblPon) _
             & " AND PIITEM=" & Val(lblItm) & " AND PIPART='')"
      clsADOCon.ExecuteSQL sSql
   End If
   PurcPRe02a.RefreshPoCursor
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set RdoEdit = Nothing
   Set AdoParameter1 = Nothing
   Set ADOParameter2 = Nothing
   Set AdoParameter3 = Nothing
   Set rdoQry = Nothing
   Set PurcPRe02c = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblMoDesc.ForeColor = vbBlack
   txtDue = Format(ES_SYSDATE, "mm/dd/yyyy")
   cmbOrigDueDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtQty = "0.000"
   txtLot = "0"
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT OPREF,PARTREF,PARTNUM FROM RnopTable," _
          & "PartTable WHERE (OPREF=PARTREF AND OPSERVPART<>'') " _
          & "ORDER BY PARTREF"
   LoadComboBox cmbMon, 1
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Grd_Click()
   Grd.Col = 0
   lblItm = Grd.Text
   Grd.Col = 1
   lblRev = Grd.Text
   Grd.Col = 0
   bItemSel = 1
   GetThisItem
   
   
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Grd.Col = 0
      lblItm = Grd.Text
      Grd.Col = 1
      lblRev = Grd.Text
      Grd.Col = 0
      bItemSel = 1
      GetThisItem
   End If
   
End Sub


Private Sub lblMoDesc_Change()
   If lblMoDesc = "*** Mo Part Is Blank ***" Then
      lblMoDesc.ForeColor = ES_RED
   Else
      lblMoDesc.ForeColor = vbBlack
   End If
End Sub

Private Sub optIns_Click()
   '
End Sub

Private Sub optIns_LostFocus()
   On Error Resume Next
   With RdoEdit
      !PIONDOCK = optIns.Value
      .Update
   End With
   
End Sub


Private Sub optSel_Click()
   'never visible-select box is open
   
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2048)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   With RdoEdit
      On Error Resume Next
      !PICOMT = txtCmt
      .Update
   End With
   CheckStatus
   
End Sub





Private Sub GetRuns()
   Dim sDescription As String
   On Error GoTo DiaErr1
   cmbRun.Clear
   bGoodMo = 0
   sSql = "SELECT DISTINCT OPRUN FROM RnopTable WHERE OPREF='" _
          & Compress(cmbMon) & "'"
   LoadNumComboBox cmbRun, "####0"
   If bSqlRows Then
      If cmbRun.ListCount > 0 Then
         bGoodMo = 1
         bGoodRun = 1
         cmbRun = cmbRun.List(0)
         GetMoPart
         cmdSel.Enabled = True
      Else
         bGoodMo = 0
         bGoodRun = 0
         cmdSel.Enabled = False
      End If
   Else
      MsgBox "No Matching Valid Runs Found."
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetMoPart()
   Dim sDescription As String
   cmbMon = GetCurrentPart(cmbMon, lblMoDesc)
   
End Sub

Private Sub AddPOItem()
   Dim iItem As Integer
   Static ItemCount As Byte
   iItem = GetNextPoItem()
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   
   sProcName = "addpoitem"
   sSql = "INSERT INTO PoitTable (PINUMBER,PIITEM,PITYPE,PIVENDOR,PIUSER,PIENTERED) " _
          & "VALUES(" & Val(lblPon) & "," & str$(iItem) & ",14,'" & Compress(PurcPRe02a.cmbVnd) & "','" & sInitials _
          & "','" & lblDte & "')"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then
      cmdNew.Enabled = False
      cmdSel.Enabled = True
      lblItm = iItem
      iRows = iRows + 1
      If iRows > 1 Then Grd.Rows = Grd.Rows + 1
      Grd.row = iRows
      Grd.Col = 3
      Grd.Text = "0.000"
      Grd.Col = 0
      Grd.Text = lblItm
      lblRev = ""
      lblServPart = ""
      lblDsc = ""
      cmbMon.Enabled = True
      cmbRun.Enabled = True
      If ItemCount = 0 Then
         If cmbMon.ListCount > 0 Then
            cmbMon = cmbMon.List(0)
            GetRuns
            ItemCount = 1
         End If
      End If
      txtQty = "0.000"
      txtPrc = "0.000"
      lblPrc = ""
      txtLot = "0"
      txtCmt = ""
      lblServPart = ""
      lblDsc = ""
      optIns.Value = vbUnchecked
      txtQty.Enabled = True
      txtPrc.Enabled = True
      txtLot.Enabled = True
      txtPrc.Enabled = True
      txtDue.Enabled = True
      cmbOrigDueDte.Enabled = True
      
      txtCmt.Enabled = True
      cmbAccount.Enabled = True
      
      optRcd.Value = vbUnchecked
      optRcd.Caption = "Received"
      cmdComments.Enabled = True
      bCantCancel = 0
      bItemSel = 1
      cmdTrm.Enabled = True
      MsgBox "The Item Has Been Successfully Added." & vbCr _
         & "Now Please Select The MO Routing Item.", _
         vbInformation, Caption
   Else
      MsgBox "Could Not Successfully Add The Item.", _
         vbInformation, Caption
   End If
   
End Sub

Private Sub FillGrid()
   Dim RdoItm As ADODB.Recordset
   
   Dim iList As Integer
   Dim sDescription As String
   iRows = 0
   Grd.Rows = 2
   MouseCursor 13
   On Error GoTo DiaErr1
   sSql = "SELECT PIITEM,PIREV,PIPART,PIPQTY FROM PoitTable WHERE " _
          & "(PITYPE <> 16 And PINUMBER =" & Val(lblPon) _
          & ") ORDER BY PIITEM,PIREV"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   If bSqlRows Then
      With RdoItm
         Do Until .EOF
            iRows = iRows + 1
            If iRows > 1 Then Grd.Rows = Grd.Rows + 1
            Grd.row = iRows
            Grd.Col = 0
            Grd.Text = Format(!PIITEM, "##0")
            Grd.Col = 1
            Grd.Text = "" & Trim(!PIREV)
            Grd.Col = 2
            Grd.Text = GetCurrentPart(!PIPART, Dummy)
            Grd.Col = 3
            Grd.Text = Format(!PIPQTY, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
      Grd.Col = 1
      lblRev = Grd.Text
      Grd.Col = 0
      lblItm = Grd.Text
      If iRows > 0 Then GetThisItem
      
   Else
      Grd.row = 1
      Grd.Col = 0
   End If
   MouseCursor 0
   Set RdoItm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub GetThisItem()
   Dim sDescription As String
   On Error GoTo DiaErr1
   'rdoQry(0) = Val(lblPon)
   'rdoQry(1) = Trim(lblRev)
   'rdoQry(2) = Val(lblItm)
   rdoQry.Parameters(0).Value = Val(lblPon)
   rdoQry.Parameters(1).Value = Trim(lblRev)
   rdoQry.Parameters(2).Value = Val(lblItm)
   bSqlRows = clsADOCon.GetQuerySet(RdoEdit, rdoQry, ES_KEYSET, True)
   If bSqlRows Then
      With RdoEdit
         txtQty = Format(!PIPQTY, ES_QuantityDataFormat)
         lblMoDesc = ""
         If Trim(!PIRUNPART) <> "" And !PIRUNNO > 0 Then
            cmbMon.Enabled = False
            cmbRun.Enabled = False
         End If
         cmbMon = "" & Trim(!PIRUNPART)
         cmbMon = GetCurrentPart(!PIRUNPART, lblMoDesc)
         sOldMoPart = cmbMon
         cmbRun = Format(!PIRUNNO, "###0")
         lblServPart = GetCurrentPart(!PIPART, lblDsc)
         txtPrc = Format(!PIESTUNIT, ES_PurchasedDataFormat)
         txtLot = Format(!PILOT, ES_QuantityDataFormat)
         If (Val(txtLot) > 0) Then
            lblPrc = Format((Val(txtPrc) * Val(txtLot)), ES_PurchasedDataFormat)
         Else
            lblPrc = ""
         End If
         txtCmt = "" & Trim(!PICOMT)
         txtDue = "" & Format(!PIPDATE, "mm/dd/yyyy")
         If Not IsNull(!PIPORIGDATE) Then
            cmbOrigDueDte = Format(!PIPORIGDATE, "mm/dd/yyyy")
        Else
            cmbOrigDueDte = ""
        End If
         
         lblOpno = Format(!PIRUNOPNO, "000")
         cmbAccount = "" & Trim(!PIACCOUNT)
         Grd.Col = 1
         Grd.Text = lblRev
         Grd.Col = 2
         Grd.Text = lblServPart
         Grd.Col = 3
         Grd.Text = txtQty
         optIns.Value = !PIONDOCK
         optDockInsted.Value = !PIONDOCKINSPECTED
         
         If !PITYPE = 15 Or !PITYPE = 17 Then
            optRcd.Value = vbChecked
            txtQty.Enabled = False
            txtPrc.Enabled = False
            txtLot.Enabled = False
            txtPrc.Enabled = False
            txtDue.Enabled = False
            cmbOrigDueDte.Enabled = False
            
            txtCmt.Enabled = False
            cmdComments.Enabled = False
            If !PITYPE = 15 Or !PITYPE = 17 Then
               bCantCancel = 1
               cmdTrm.Enabled = False
               If !PITYPE = 17 Then optRcd.Caption = "Invoiced" Else optRcd.Caption = "Received"
               cmdTrm.Enabled = False
            Else
               If !PITYPE = 14 Then
                  bCantCancel = 0
                  cmdTrm.Enabled = True
                  optRcd.Caption = "Received"
                  cmdTrm.Enabled = True
               End If
            End If
         Else
            txtQty.Enabled = True
            txtPrc.Enabled = True
            txtLot.Enabled = True
            txtPrc.Enabled = True
            txtDue.Enabled = True
            cmbOrigDueDte.Enabled = True
            txtCmt.Enabled = True
            optRcd.Value = vbUnchecked
            optRcd.Caption = "Received"
            cmdComments.Enabled = True
            bCantCancel = 0
            cmdTrm.Enabled = True
         End If
         '5/19/04
         If bAllowPrice = 1 And !PITYPE = 15 Then txtPrc.Enabled = True
      End With
      Grd.Col = 0
   End If
   bOnLoad = 0
   CheckStatus
   Exit Sub
   
DiaErr1:
   sProcName = "gethisitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
Private Sub txtDue_DropDown()
   ShowCalendarEx Me
   bFieldChanged = 1
End Sub

Private Sub txtDue_GotFocus()
   bFieldChanged = 0
End Sub

Private Sub txtDue_LostFocus()
   If Len(txtDue) > 0 Then txtDue = CheckDateEx(txtDue)
   With RdoEdit
      On Error Resume Next
      !PIPDATE = Format(txtDue, "mm/dd/yyyy")
      .Update
        
   End With
   CheckStatus
   
   If Len(txtDue) = 0 And bFieldChanged = 1 Then bFieldChanged = 0


   If (SysCalendar.Visible = False) Then
      With RdoEdit
         On Error Resume Next
         If MsgBox("Would you also like to change the original due date to " & txtDue & " ?", vbYesNo, Caption) = vbYes Then
            !PIPORIGDATE = Format(txtDue, "mm/dd/yyyy")
            .Update
            cmbOrigDueDte = Format(txtDue, "mm/dd/yyyy")
         End If

      End With
   End If
   
   
End Sub

Private Sub cmbOrigDueDte_DropDown() 'BBS Added on 03/09/2010 for Ticket #11364
  ShowCalendarEx Me
   bFieldChanged = 1
End Sub

Private Sub cmbOrigDueDte_LostFocus()  'BBS Added on 03/09/2010 for Ticket #11364
   cmbOrigDueDte = CheckDateEx(cmbOrigDueDte)
'   If Val(txtQty) > 0 Then
'      cmdNew.Enabled = True
'      If bCantCancel = 0 Then cmdTrm.Enabled = True
'   End If
   With RdoEdit
      !PIPORIGDATE = cmbOrigDueDte
      .Update
   End With
End Sub

Private Sub cmbOrigDueDte_GotFocus()
   bFieldChanged = 0
End Sub

Private Sub txtlot_LostFocus()
   txtLot = CheckLen(txtLot, 9)
   txtLot = Format(Abs(Val(txtLot)), ES_QuantityDataFormat)
   With RdoEdit
      On Error Resume Next
      If (txtLot <> "" And Val(txtLot) > 0) Then
        SetPriceAsUnitCost False
      Else
        lblPrc = ""
      End If
      
      !PILOT = Val(txtLot)
      .Update
   End With
   
End Sub


Private Sub txtPrc_LostFocus()
   txtPrc = CheckLen(txtPrc, 9)
   txtPrc = Format(Abs(Val(txtPrc)), ES_PurchasedDataFormat)
      On Error Resume Next
      If (txtLot <> "" And Val(txtLot) > 0) Then
        SetPriceAsUnitCost True
      Else
        With RdoEdit
         !PIESTUNIT = Format(Val(txtPrc), ES_PurchasedDataFormat)
         If !PIAMT > 0 Or (!PITYPE >= 15) Then
              !PIAMT = Format(Val(txtPrc), ES_PurchasedDataFormat)
         End If
         .Update
        End With
            ' update the text
            ' MM lblPrc = txtPrc
        
     End If
   CheckStatus
   
End Sub


Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 10)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   With RdoEdit
      On Error Resume Next
        !PIPQTY = Format(Val(txtQty), ES_QuantityDataFormat)
        .Update
   End With
   CheckStatus
   Grd.Col = 3
   Grd.Text = txtQty
   Grd.Col = 0
   
End Sub



Private Function GetNextPoItem() As Integer
   Dim RdoNxt As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT MAX(PIITEM) FROM PoitTable WHERE " _
          & "PINUMBER=" & Val(lblPon) & " "
   'Set RdoNxt = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   Set RdoNxt = clsADOCon.GetRecordSet(sSql)
   If Not IsNull(RdoNxt.Fields(0)) Then
      If Not IsNull(RdoNxt.Fields(0)) Then
         If Val(RdoNxt.Fields(0)) > 0 Then GetNextPoItem = RdoNxt.Fields(0) + 1
      Else
         GetNextPoItem = 1
      End If
   Else
      GetNextPoItem = 1
   End If
   Set RdoNxt = Nothing
   Exit Function
   
DiaErr1:
   GetNextPoItem = 1
   
End Function

Private Sub CheckStatus()
   If Len(Trim(cmbMon)) > 0 And Len(Trim(lblServPart)) > 0 And _
          Val(cmbRun) > 0 Then cmdNew.Enabled = True
   
End Sub

Private Function SetPriceAsUnitCost(bFromPrice As Boolean)
    
    Dim strMsg As String
    Dim bResponse As Byte
    
    strMsg = "Is the Price a Lot Charge?"
    bResponse = MsgBox(strMsg, ES_YESQUESTION, Caption)
    
    If (bResponse = vbYes) Then
        lblPrc = txtPrc
        txtPrc = Format((Val(txtPrc) / Val(txtLot)), ES_PurchasedDataFormat)
        With RdoEdit
         !PIESTUNIT = Format(Val(txtPrc), ES_PurchasedDataFormat)
         .Update
        End With
        txtPrc.ToolTipText = "Price as Unit cost"
    Else
        If (bFromPrice) Then
            lblPrc = ""
            txtLot = Format(0, ES_QuantityDataFormat)
            With RdoEdit
             !PILOT = Val(txtLot)
             .Update
            End With
        End If
    End If
End Function


Private Sub FillAccounts()
   Dim rdo As ADODB.Recordset
   
   sSql = "Qry_FillLowAccounts"
   LoadComboBox cmbAccount
   
   'set the po acct to the default
   sSql = "SELECT rtrim(COAPACCT) FROM ComnTable WHERE COREF=1"
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      If Len(rdo.Fields(0)) > 0 Then
         'On Error Resume Next
         cmbAccount = rdo.Fields(0)
         Err.Clear
      End If
   End If
   Set rdo = Nothing
End Sub




Private Function CheckPoAccounts() As Byte
   'On Error Resume Next
   Dim RdoChk As ADODB.Recordset
   
   CheckPoAccounts = 0
   sSql = "SELECT isnull(PurchaseAccount, 0) FROM Preferences WHERE " _
          & "PreRecord=1"
   If clsADOCon.GetDataSet(sSql, RdoChk, ES_FORWARD) Then
      CheckPoAccounts = RdoChk.Fields(0)
   End If
   Set RdoChk = Nothing
End Function

Private Function IsValidMONumber(sMO As String) As Boolean
    Dim RdoMO As ADODB.Recordset
    
    sSql = "SELECT TOP 1 OPREF FROM RnopTable WHERE OPREF = '" & Compress(sMO) & "' "
    IsValidMONumber = clsADOCon.GetDataSet(sSql, RdoMO, ES_FORWARD)
    Set RdoMO = Nothing
End Function
