VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routings"
   ClientHeight    =   6090
   ClientLeft      =   1845
   ClientTop       =   1290
   ClientWidth     =   6420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3101
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6090
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkInactive 
      Height          =   255
      Left            =   1500
      TabIndex        =   48
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdAssnPic 
      Caption         =   "&Assign Pic"
      Height          =   285
      Left            =   5400
      TabIndex        =   47
      ToolTipText     =   "Assign picture to routing."
      Top             =   2280
      Width           =   940
   End
   Begin VB.CheckBox chkUpChild 
      Caption         =   "Check1"
      Height          =   255
      Left            =   5040
      TabIndex        =   46
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame z2 
      Height          =   30
      Left            =   120
      TabIndex        =   45
      Top             =   1320
      Width           =   6252
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtCmt 
      Height          =   1185
      Left            =   1500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Tag             =   "9"
      Text            =   "RoutRTe01a.frx":07AE
      ToolTipText     =   "Comment (5120 Chars Max)"
      Top             =   3720
      Width           =   4335
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   285
      Left            =   5400
      TabIndex        =   37
      ToolTipText     =   "Update Parts Lists To Current Labor Estimates"
      Top             =   5400
      Width           =   940
   End
   Begin VB.TextBox txtQdy 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Acculated Hours"
      Top             =   5040
      Width           =   825
   End
   Begin VB.TextBox txtMdy 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Acculated Hours"
      Top             =   5040
      Width           =   825
   End
   Begin VB.TextBox txtSet 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Acculated Hours"
      Top             =   5400
      Width           =   825
   End
   Begin VB.TextBox txtUnt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Acculated Hours"
      Top             =   5400
      Width           =   825
   End
   Begin VB.ComboBox txtRby 
      Height          =   315
      Left            =   1500
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "Engineer"
      Top             =   1440
      Width           =   2340
   End
   Begin VB.TextBox txtApp 
      Height          =   285
      Left            =   1500
      TabIndex        =   4
      Tag             =   "2"
      ToolTipText     =   "Approval Name"
      Top             =   1800
      Width           =   2085
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "RoutRTe01a.frx":07B5
      Height          =   320
      Left            =   4920
      Picture         =   "RoutRTe01a.frx":1127
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Parts Assigned To This Routing"
      Top             =   640
      Width           =   350
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1440
      Width           =   1250
   End
   Begin VB.ComboBox txtAte 
      Height          =   315
      Left            =   5160
      TabIndex        =   5
      Tag             =   "4"
      Top             =   1800
      Width           =   1250
   End
   Begin VB.CheckBox optSet 
      Caption         =   "Check1"
      Height          =   255
      Left            =   5040
      TabIndex        =   24
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Settings"
      Height          =   285
      Left            =   5400
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Set Routing Defaults"
      Top             =   960
      Width           =   940
   End
   Begin VB.CheckBox optOps 
      Height          =   255
      Left            =   5400
      TabIndex        =   22
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "&Assign"
      Height          =   285
      Left            =   5400
      TabIndex        =   8
      ToolTipText     =   "Assign This Routing To Parts"
      Top             =   3000
      Width           =   940
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1500
      Sorted          =   -1  'True
      TabIndex        =   7
      Tag             =   "3"
      ToolTipText     =   "Select Part Number To Assign"
      Top             =   2610
      Width           =   3345
   End
   Begin VB.TextBox txtRev 
      Height          =   285
      Left            =   1500
      TabIndex        =   6
      Tag             =   "3"
      ToolTipText     =   "Revision Of Routing"
      Top             =   2250
      Width           =   465
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "&Operations"
      Height          =   285
      Left            =   5400
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Show Operations"
      Top             =   600
      Width           =   940
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1500
      TabIndex        =   1
      Tag             =   "2"
      Text            =   " "
      ToolTipText     =   "(30) Char Maximun"
      Top             =   990
      Width           =   3075
   End
   Begin VB.ComboBox cmbRte 
      Height          =   288
      Left            =   1500
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Add/Edit Routing"
      Top             =   630
      WhatsThisHelpID =   100
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5400
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   940
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   6120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6090
      FormDesignWidth =   6420
   End
   Begin VB.Label Label1 
      Caption         =   "Inactive"
      Height          =   255
      Left            =   180
      TabIndex        =   49
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision Notes"
      Height          =   285
      Index           =   14
      Left            =   180
      TabIndex        =   43
      Top             =   3675
      Width           =   1425
   End
   Begin VB.Label lblRunQty 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   42
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hrs"
      Height          =   285
      Index           =   18
      Left            =   4920
      TabIndex        =   41
      Top             =   4200
      Width           =   465
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hrs"
      Height          =   285
      Index           =   17
      Left            =   2280
      TabIndex        =   40
      Top             =   5400
      Width           =   465
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hrs"
      Height          =   285
      Index           =   16
      Left            =   4920
      TabIndex        =   39
      Top             =   5040
      Width           =   465
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hrs"
      Height          =   285
      Index           =   15
      Left            =   2280
      TabIndex        =   38
      Top             =   5040
      Width           =   465
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Queue"
      Height          =   285
      Index           =   13
      Left            =   180
      TabIndex        =   36
      Top             =   5040
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Move"
      Height          =   285
      Index           =   12
      Left            =   2880
      TabIndex        =   35
      Top             =   5040
      Width           =   885
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Setup "
      Height          =   285
      Index           =   11
      Left            =   180
      TabIndex        =   34
      Top             =   5400
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Unit/Cy"
      Height          =   285
      Index           =   10
      Left            =   2880
      TabIndex        =   33
      Top             =   5400
      Width           =   1485
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   285
      Index           =   9
      Left            =   4920
      TabIndex        =   27
      ToolTipText     =   "Routing Date As 08/08/97,08 08 97 or 08-08-97"
      Top             =   2610
      Width           =   615
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5520
      TabIndex        =   26
      Top             =   2610
      Width           =   405
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1500
      TabIndex        =   25
      Top             =   2960
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Routing"
      Height          =   285
      Index           =   8
      Left            =   180
      TabIndex        =   21
      Top             =   3330
      Width           =   1425
   End
   Begin VB.Label lblRout 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1500
      TabIndex        =   20
      Top             =   3330
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   7
      Left            =   180
      TabIndex        =   19
      Top             =   2610
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   285
      Index           =   6
      Left            =   180
      TabIndex        =   18
      Top             =   2250
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "App Date"
      Height          =   285
      Index           =   5
      Left            =   4200
      TabIndex        =   17
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Approved By"
      Height          =   285
      Index           =   4
      Left            =   180
      TabIndex        =   16
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   285
      Index           =   3
      Left            =   4200
      TabIndex        =   15
      ToolTipText     =   "Routing Date As 08/08/97,08 08 97 or 08-08-97"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing By"
      Height          =   285
      Index           =   2
      Left            =   180
      TabIndex        =   14
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   13
      Top             =   990
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add/Revise Routing"
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   12
      Top             =   540
      Width           =   915
   End
End
Attribute VB_Name = "RoutRTe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'9/29/04 added Operation Hours
'10/13/04 added Calculate BOM Estimate
'5/10/05 added RTREVNOTES
'12/2/05 Removed RdoRes reference
Option Explicit
'Dim RdoStm As rdoQuery
'Dim RdoRte As rdoQuery
'Dim RdoRtg As ADODB.Recordset

Dim AdoCmdStm As ADODB.Command
Dim AdoCmdRte As ADODB.Command
Dim RdoRtg As ADODB.Recordset


Dim bCanceled As Byte
Dim bGoodRout As Byte
Dim bOnLoad As Byte
Dim bNewRout As Byte
Dim bRoutSec As Byte

Dim strPrevREv As String

Dim lEstRun As Long

Private txtKeyPress() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtQdy.BackColor = Es_TextDisabled
   txtMdy.BackColor = Es_TextDisabled
   txtSet.BackColor = Es_TextDisabled
   txtUnt.BackColor = Es_TextDisabled
   lblRunQty = "0"
   
End Sub



Private Sub chkInactive_LostFocus()
   If bGoodRout Then
      On Error Resume Next
      'RdoRtg.Edit
      RdoRtg!RTINACTIVE = "" & chkInactive
      RdoRtg.Update
      If Err > 0 Then ValidateEdit
   End If

End Sub

Private Sub cmbPrt_Click()
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   
End Sub


Private Sub cmbRte_Click()
   bGoodRout = GetRout(True)
   
End Sub

Private Sub cmbRte_LostFocus()
   cmbRte = CheckLen(cmbRte, 30)
   If bCanceled Then Exit Sub
   If Len(cmbRte) = 0 Then
      On Error Resume Next
      bGoodRout = False
      cmdOpt.Enabled = False
      cmdCan.SetFocus
      Exit Sub
   Else
      bGoodRout = GetRout(True)
   End If
   If Not bGoodRout Then AddRouting
   
End Sub

Private Sub cmdAsn_Click()
   If cmbPrt <> "NONE" Then
      On Error Resume Next
      MouseCursor 11
      clsADOCon.ExecuteSQL "UPDATE PartTable SET PAROUTING='" & Compress(RdoRtg!RTREF) & "' WHERE PARTNUM='" & cmbPrt & "'"
      MouseCursor 0
      If clsADOCon.RowsAffected > 0 Then
         lblRout = cmbRte
         SysMsg "Routing Assigned", True, Me
      Else
         MsgBox "Part Wasn't Found.", vbInformation, Caption
      End If
   End If
   
End Sub

Private Sub cmdAssnPic_Click()
   RoutRTe01g.txtRte = cmbRte.Text
   RoutRTe01g.txtDsc = txtDsc.Text
   RoutRTe01g.Show
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3101
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdOpt_Click()
   MouseCursor 13
   bGoodRout = GetRout(True)
   If bGoodRout Then
      
      If ((bRoutSec = True) And (Trim(txtApp) <> "")) Then
         ' Force Revision change
         Dim strRev As String
         strRev = txtRev.Text
      
         RevisionChange strRev, True
         If (chkUpChild = 0) Then
            MsgBox "Please change the revision number to modify the Operations.", vbCritical
            Exit Sub
         Else
            'RdoRtg.Edit
            RdoRtg!RTREV = "" & txtRev
            RdoRtg!RTBY = "" & txtRby
            If Len(txtDte) > 0 Then
               RdoRtg!RTDATE = Format(txtDte, "mm/dd/yyyy")
            Else
               RdoRtg!RTDATE = Null
            End If
            RdoRtg!RTREVNOTES = "" & txtCmt
            
            RdoRtg.Update
            ' Clear the approval names
            ResetApproval
         End If
      End If
      
      optOps.Value = vbChecked
      sPassedRout = "" & Trim(RdoRtg!RTREF)
      RoutRTe01b.lblRout = "" & Trim(RdoRtg!RTREF)
      RoutRTe01b.Caption = RoutRTe01b.Caption & " For " & cmbRte
      RoutRTe01b.Show
   End If
   
End Sub

Private Sub cmdSet_Click()
   optSet.Value = vbChecked
   RoutRTe01c.Show
   
End Sub

Private Sub cmdUpd_Click()
   On Error Resume Next
   RoutRTe01d.Show vbModal
   lEstRun = Val(lblRunQty)
   If lEstRun = 0 Then
      MsgBox "Requires A Run Quantity.", _
         vbInformation, Caption
   Else
      CalculateBomEst
   End If
   lblRunQty = "0"
   
End Sub

Private Sub cmdVew_Click()
   If cmdVew Then
      RteTree.Show
      cmdVew = False
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillParts
      FillRoutings
      If Len(sCurrRout) Then cmbRte = sCurrRout 'bbs re-added this on 3/21/2016
      bNewRout = False
      bGoodRout = GetRout(False)
      bRoutSec = GetRoutSecurity()
   End If
   If optOps.Value = vbChecked Then
      optOps.Value = vbUnchecked
      Unload RoutRTe01b
   End If
   If optSet.Value = vbChecked Then
      optSet.Value = vbUnchecked
      Unload RoutRTe01c
   End If
   chkUpChild = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   On Error Resume Next
   sSql = "SELECT * FROM RthdTable WHERE RTREF= ? "
   
   Set AdoCmdStm = New ADODB.Command
   AdoCmdStm.CommandText = sSql
   
   Dim prmPrtRef As ADODB.Parameter
   Set prmPrtRef = New ADODB.Parameter
   prmPrtRef.Type = adChar
   prmPrtRef.Size = 30
   AdoCmdStm.Parameters.Append prmPrtRef
   
   'Set RdoStm = RdoCon.CreateQuery("", sSql)
   ' TODO: RdoStm.MaxRows = 1
   
   sSql = "SELECT DISTINCT PARTNUM,PADESC,PAROUTING FROM PartTable WHERE PAROUTING= ? "
   
   Set AdoCmdRte = New ADODB.Command
   AdoCmdRte.CommandText = sSql
   
   Dim prmPARte As ADODB.Parameter
   Set prmPARte = New ADODB.Parameter
   prmPARte.Type = adChar
   prmPARte.Size = 30
   AdoCmdRte.Parameters.Append prmPARte
   
   
   'Set RdoRte = RdoCon.CreateQuery("", sSql)
   'RdoStm.MaxRows = 1
   bOnLoad = 1
   GetRoutingIncrementDefault
   
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If optOps.Value = vbChecked Then Unload RoutRTe01b
   If bGoodRout Then
      sCurrRout = cmbRte
      SaveSetting "Esi2000", "EsiEngr", "CurrentRouting", Trim(sCurrRout)
   Else
      sCurrRout = ""
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoCmdStm = Nothing
   Set AdoCmdRte = Nothing
   Set RdoRtg = Nothing
   FormUnload
   Set RoutRTe01a = Nothing
   
End Sub




Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optOps_Click()
   'never visible
   'used to see if BompBM02b should be unloaded
   
End Sub

Private Sub optSet_Click()
   'never visible-Is Select Open?
End Sub

Private Sub txtApp_LostFocus()
   txtApp = CheckLen(txtApp, 20)
   txtApp = StrCase(txtApp)
   If bGoodRout Then
      On Error Resume Next
      'RdoRtg.Edit
      RdoRtg!RTAPPBY = "" & txtApp
      RdoRtg.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtAte_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtAte_LostFocus()
   If Trim(txtAte) = "" Then txtAte = CheckDateEx(txtAte)
   If bGoodRout Then
      On Error Resume Next
      'RdoRtg.Edit
      If Len(txtAte) > 0 Then
         RdoRtg!RTAPPDATE = Format(txtAte, "mm/dd/yyyy")
      Else
         RdoRtg!RTAPPDATE = Null
      End If
      RdoRtg.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 1020)
   If bGoodRout Then
      On Error Resume Next
      'RdoRtg.Edit
      RdoRtg!RTREVNOTES = "" & txtCmt
      RdoRtg.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   If bGoodRout Then
      On Error Resume Next
      'RdoRtg.Edit
      RdoRtg!RTDESC = "" & txtDsc
      RdoRtg.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtDte_LostFocus()
   If Trim(txtDte) = "" Then txtDte = CheckDateEx(txtDte)
   If bGoodRout Then
      On Error Resume Next
      'RdoRtg.Edit
      If Len(txtDte) > 0 Then
         RdoRtg!RTDATE = Format(txtDte, "mm/dd/yyyy")
      Else
         RdoRtg!RTDATE = Null
      End If
      RdoRtg.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtRby_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   txtRby = CheckLen(txtRby, 20)
   txtRby = StrCase(txtRby)
   If bGoodRout Then
      On Error Resume Next
      'RdoRtg.Edit
      RdoRtg!RTBY = "" & txtRby
      RdoRtg.Update
      If Err > 0 Then ValidateEdit
   End If
   For iList = 0 To txtRby.ListCount - 1
      If UCase$(txtRby) = UCase$(txtRby.List(iList)) Then b = 1
   Next
   If b = 0 Then txtRby.AddItem txtRby
   
End Sub


Private Sub txtRev_GotFocus()
    strPrevREv = txtRev.Text
End Sub

Private Sub txtRev_LostFocus()
   Dim bres As Boolean
   txtRev = CheckLen(txtRev, 2)
   If bGoodRout Then
      On Error Resume Next
      
      If strPrevREv <> txtRev.Text Then
      
         If (bRoutSec = True) Then
            
            RevisionChange txtRev.Text, False
            
            If (chkUpChild = 0) Then
               MsgBox "Please change the revision number to modify the Operations.", vbCritical
               Exit Sub
            Else
               'RdoRtg.Edit
               RdoRtg!RTREV = "" & txtRev
               RdoRtg!RTBY = "" & txtRby
               If Len(txtDte) > 0 Then
                  RdoRtg!RTDATE = Format(txtDte, "mm/dd/yyyy")
               Else
                  RdoRtg!RTDATE = Null
               End If
               RdoRtg!RTREVNOTES = "" & txtCmt
               RdoRtg.Update
            End If
            ResetApproval
         Else
            'RdoRtg.Edit
            RdoRtg!RTREV = "" & txtRev
            RdoRtg.Update
            ResetApproval
         End If
      End If
      
      If Err > 0 Then ValidateEdit
   End If
End Sub



Private Function GetRout(bOpen As Byte) As Byte
   Dim RdoPrt As ADODB.Recordset
   GetRout = False
   On Error GoTo DiaErr1
   ' TODO RdoStm.RowsetSize = 1
   AdoCmdStm.Parameters(0).Value = Compress(cmbRte)
   'RdoStm(0) = Compress(cmbRte)
   
   ' NOT SURE RdoRtg.MaxRecords = 1
   bSqlRows = clsADOCon.GetQuerySet(RdoRtg, AdoCmdStm, ES_KEYSET, True, 1)
   
   If bSqlRows Then
      With RdoRtg
         GetRout = True
         cmbRte = "" & Trim(!RTNUM)
         If bNewRout = 0 Then
            txtDsc = "" & Trim(!RTDESC)
            txtRby = "" & Trim(!RTBY)
            txtDte = "" & Format(!RTDATE, "mm/dd/yyyy")
            txtApp = "" & Trim(!RTAPPBY)
            txtAte = "" & Format(!RTAPPDATE, "mm/dd/yyyy")
            txtRev = "" & Trim(!RTREV)
            txtCmt = "" & Trim(!RTREVNOTES)
            txtQdy = Format(!RTQUEUEHRS, ES_QuantityDataFormat)
            txtMdy = Format(!RTMOVEHRS, ES_QuantityDataFormat)
            txtSet = Format(!RTSETUPHRS, ES_QuantityDataFormat)
            txtUnt = Format(!RTUNITHRS, ES_QuantityDataFormat)
            lblRout = "" & Trim(!RTNUM)
            chkInactive = !RTINACTIVE
         Else
            txtDsc = " "
            txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
            txtApp = " "
            txtAte = "" & Format(!RTAPPDATE, "mm/dd/yyyy")
            txtRev = " "
            txtCmt = ""
            ' Now set the new route as zero
            bNewRout = 0
            chkInactive = 0
         End If
         
         ' Set prev Revision as
         strPrevREv = txtRev.Text
         
         cmdOpt.Enabled = True
         If bOpen Then
            txtRby.Enabled = True
            txtDte.Enabled = True
            txtApp.Enabled = True
            txtAte.Enabled = True
            txtRev.Enabled = True
            txtCmt.Enabled = True
            cmbPrt.Enabled = True
            cmdAsn.Enabled = True
            cmdUpd.Enabled = True
            chkInactive.Enabled = True
         Else
            txtRby.Enabled = False
            txtDte.Enabled = False
            txtRev.Enabled = False
            txtApp.Enabled = False
            txtAte.Enabled = False
            txtCmt.Enabled = False
            cmbPrt.Enabled = False
            cmdAsn.Enabled = False
            cmdUpd.Enabled = False
            chkInactive.Enabled = False
         End If
      End With
      GetRout = True
   Else
      txtDsc = ""
      txtRby = ""
      txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
      txtApp = ""
      txtAte = ""
      txtRev = ""
      chkInactive = 0
      GetRout = False
      txtRby.Enabled = False
      txtDte.Enabled = False
      txtRev.Enabled = False
      txtApp.Enabled = False
      txtAte.Enabled = False
      cmbPrt.Enabled = False
      cmdOpt.Enabled = False
      cmdAsn.Enabled = False
      cmdUpd.Enabled = False
      chkInactive.Enabled = False
    
      On Error Resume Next
      cmbRte.SetFocus
      Exit Function
   End If
   
   'TODO RdoRte.RowsetSize = 1
   'RdoRte(0) = Compress(cmbRte)
   AdoCmdRte.Parameters(0).Value = Compress(cmbRte)
   ' MAY BE   RdoPrt.MaxRecords = 1
   bSqlRows = clsADOCon.GetQuerySet(RdoPrt, AdoCmdRte, ES_KEYSET, True, 1)
   If bSqlRows Then
      cmbPrt = "" & Trim(RdoPrt!PartNum)
      lblDsc = "" & Trim(RdoPrt!PADESC)
      If Val(txtUnt) = 0 Then GetOperationTimes
   Else
      cmbPrt = ""
      lblDsc = ""
   End If
   On Error Resume Next
   RdoPrt.Close
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrout"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddRouting()
   Dim sNewRout As String
   Dim bResponse As Byte
   
   bResponse = MsgBox(cmbRte & " Wasn't Found. Add It?", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      bGoodRout = False
      On Error Resume Next
      cmbRte = cmbRte.List(0)
      cmbRte.SetFocus
      GetRout (False)
      Width = Width + 10
      Exit Sub
   End If
   bResponse = IllegalCharacters(cmbRte)
   If bResponse > 0 Then
      MsgBox "Routing Number Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   MouseCursor 11
   sNewRout = Compress(cmbRte)
   On Error GoTo RoutRTeAdd1
   sSql = "INSERT INTO RthdTable (RTREF,RTNUM) VALUES('" & sNewRout & "','" _
          & cmbRte & "')"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   On Error Resume Next
   AddComboStr cmbRte.hwnd, cmbRte
   MouseCursor 0
   bNewRout = 1
   bGoodRout = GetRout(True)
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtDsc.SetFocus
   SysMsg cmbRte & " Added.", True, Me
   Exit Sub
   
RoutRTeAdd1:
   sProcName = "addrouting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume RoutRTeAdd2
RoutRTeAdd2:
   MouseCursor 0
   On Error Resume Next
   clsADOCon.RollbackTrans
   RdoRtg.Close
   MsgBox CurrError.Description & vbCrLf & "Couldn't Add Routing.", vbExclamation, Caption
   DoModuleErrors Me
   
End Sub



Private Sub FillParts()
   Dim RdoPrt As ADODB.Recordset
   cmbPrt.Clear
   On Error GoTo DiaErr1
   sSql = "Qry_FillPartRoutings"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblRout = "" & Trim(!PAROUTING)
         Do Until .EOF
            AddComboStr cmbPrt.hwnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoPrt
      End With
   Else
      lblRout = ""
      lblDsc = ""
   End If
   sSql = "SELECT DISTINCT RTBY FROM RthdTable ORDER BY RTBY "
   LoadComboBox txtRby, -1
   If txtRby.ListCount > 0 Then txtRby = txtRby.List(0)
   Set RdoPrt = Nothing
   If lblRout = "" Then
      lblRout = "NONE"
   Else
      GetPart
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetPart()
   Dim RdoPrt As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetPartRouting '" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblRout = "" & Trim(!PAROUTING)
         lblDsc = "" & Trim(!PADESC)
         lblTyp = Format(0 + !PALEVEL, "0")
      End With
   Else
      lblRout = ""
      lblDsc = ""
   End If
   If lblRout = "" Then
      lblRout = "NONE"
      RdoPrt.Close
      Exit Sub
   End If
   sSql = "Qry_GetRoutingBasics '" & lblRout & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_STATIC)
   If bSqlRows Then lblRout = "" & Trim(RdoPrt!RTNUM)
   Set RdoPrt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetRoutSecurity()
   On Error GoTo DiaErr1
   Dim RdoRout As ADODB.Recordset
   Err = 0
   sSql = "SELECT ISNULL(COROUTSEC, 0) COROUTSEC FROM ComnTable WHERE COREF=1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRout, ES_FORWARD)
   If bSqlRows Then
      With RdoRout
         GetRoutSecurity = IIf((!COROUTSEC = 0), False, True)
         ClearResultSet RdoRout
      End With
   Else
      GetRoutSecurity = False
   End If
   Set RdoRout = Nothing
   Exit Function
   
   
DiaErr1:
   sProcName = "GetRoutSecurity"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Public Sub GetOperationTimes()
   Dim RdoHrs As ADODB.Recordset
   Dim cSetup As Currency
   Dim cUnit As Currency
   Dim cQueue As Currency
   Dim cMove As Currency
   
   On Error GoTo DiaErr1
   Err = 0
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT SUM(OPSETUP),SUM(OPUNIT),SUM(OPQHRS)," _
          & "SUM(OPMHRS) FROM RtopTable WHERE OPREF='" _
          & Compress(cmbRte) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoHrs, ES_FORWARD)
   If bSqlRows Then
      With RdoHrs
         cSetup = Format(.Fields(0), ES_QuantityDataFormat)
         cUnit = Format(.Fields(1), ES_QuantityDataFormat)
         cQueue = Format(.Fields(2), ES_QuantityDataFormat)
         cMove = Format(.Fields(3), ES_QuantityDataFormat)
         ClearResultSet RdoHrs
      End With
   Else
      cSetup = 0
      cUnit = 0
      cQueue = 0
      cMove = 0
   End If
   txtQdy = Format(cQueue, ES_QuantityDataFormat)
   txtMdy = Format(cMove, ES_QuantityDataFormat)
   txtSet = Format(cSetup, ES_QuantityDataFormat)
   txtUnt = Format(cUnit, ES_QuantityDataFormat)
   If bGoodRout Then
      If clsADOCon.ADOErrNum = 0 Then
         'RdoRtg.Edit
         RdoRtg!RTQUEUEHRS = cQueue
         RdoRtg!RTMOVEHRS = cMove
         RdoRtg!RTSETUPHRS = cSetup
         RdoRtg!RTUNITHRS = cUnit
         RdoRtg.Update
      End If
   End If
   Set RdoHrs = Nothing
   Exit Sub
DiaErr1:
   txtQdy = "0.000"
   txtMdy = "0.000"
   txtSet = "0.000"
   txtUnt = "0.000"
   
End Sub

Private Function ResetApproval()
    If bGoodRout Then
        txtApp = ""
        txtAte = ""
        On Error Resume Next
        'RdoRtg.Edit
        RdoRtg!RTAPPBY = "" & txtApp
        RdoRtg!RTAPPDATE = Null
        RdoRtg.Update
        If Err > 0 Then ValidateEdit
    End If
End Function
'10/13/04 New

Private Function RevisionChange(strRev As String, bIncRev As Boolean)
   
   Dim strNewRev As String
   strNewRev = strRev
   If (Trim(strNewRev) = "") Then strNewRev = "0"

   If (bIncRev = True) Then
      If (IsNumeric(strNewRev)) Then
         strNewRev = CStr(CDbl(strNewRev) + 1)
      Else
         strNewRev = Chr$(Asc(strNewRev) + 1)
      End If
   End If
   
   RoutRTe01f.txtRev = strNewRev
   RoutRTe01f.txtRby = txtRby
   RoutRTe01f.txtDte = txtDte
   
   ' We don't need to initilize
   RoutRTe01f.txtCmt = "" 'txtCmt
   
   RoutRTe01f.txtPrevRev = strPrevREv
   RoutRTe01f.txtRtByPrev = txtRby
   
   chkUpChild = 0
   RoutRTe01f.Show vbModal

End Function

Public Sub CalculateBomEst()
   Dim RdoPar As ADODB.Recordset
   Dim bResponse As Byte
   Dim iCounter As Integer
   Dim iRow As Integer
   Dim sMsg As String
   Dim cCurLab
   Dim cCurLabOh As Currency
   Dim cEstLab As Currency
   Dim cEstLabOh As Currency
   Dim sPartNumbers() As String
   
   On Error GoTo DiaErr1
   sMsg = "You Are About To Update Parts Lists Labor For " & vbCrLf _
          & "A Quantity Of " & lEstRun & " . Continue."
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
   Else
      sSql = "select PARTREF,PAROUTING from PartTable WHERE " _
             & "PAROUTING='" & Compress(cmbRte) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPar, ES_FORWARD)
      If bSqlRows Then
         With RdoPar
            Do Until .EOF
               iCounter = iCounter + 1
               ReDim Preserve sPartNumbers(iCounter)
               sPartNumbers(iCounter) = "" & Trim(!PartRef)
               .MoveNext
            Loop
            ClearResultSet RdoPar
         End With
      End If
      If iCounter = 0 Then
         MsgBox "No Part Numbers Found With This Routing.", _
            vbInformation, Caption
         Exit Sub
      Else
         sSql = "SELECT OPREF,OPSETUP,OPUNIT,OPSHOP,OPCENTER," _
                & "WCNREF,WCNSHOP,WCNOHPCT,WCNSTDRATE FROM RtopTable,WcntTable " _
                & "WHERE (OPSHOP=WCNSHOP AND OPCENTER=WCNREF AND OPREF='" _
                & Compress(cmbRte) & "') AND (OPSETUP+OPUNIT>0)"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoPar, ES_FORWARD)
         If bSqlRows Then
            With RdoPar
               Do Until .EOF
                  cCurLab = ((!OPUNIT * lEstRun) * !WCNSTDRATE)
                  cCurLab = cCurLab + (!OPSETUP * !WCNSTDRATE)
                  cCurLabOh = (!WCNOHPCT / 100)
                  cCurLabOh = cCurLabOh * cCurLab
                  cEstLab = cEstLab + cCurLab
                  cEstLabOh = cEstLabOh + cCurLabOh
                  .MoveNext
               Loop
               ClearResultSet RdoPar
            End With
            If cEstLab <= 0 Then cEstLabOh = 0 _
                          Else cEstLabOh = cEstLabOh / cEstLab
            cEstLabOh = cEstLabOh * 100
            clsADOCon.BeginTrans
            On Error Resume Next
            For iRow = 1 To iCounter
               sSql = "UPDATE BmplTable SET BMESTLABOR=" & cEstLab _
                      & ",BMESTLABOROH=" & cEstLabOh & " WHERE " _
                      & "BMPARTREF='" & sPartNumbers(iRow) & "'"
               clsADOCon.ExecuteSQL sSql 'rdExecDirect
            Next
            sMsg = "Unit Estimated Labor Cost Is " & Format(cEstLab, ES_QuantityDataFormat) & vbCrLf _
                   & "Estimated Labor Overhead Is " & Format(cEstLabOh, "##0.00") & "%" & vbCrLf _
                   & "Update The " & iCounter & " Matching Parts List Rows?"
            bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
            If bResponse = vbYes Then
               clsADOCon.CommitTrans
               SysMsg "All Available Rows Were Updated.", True
            Else
               clsADOCon.RollbackTrans
               CancelTrans
            End If
         End If
      End If
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "calculatebome"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

