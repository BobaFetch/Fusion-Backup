VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaSCe01a
   BorderStyle = 3 'Fixed Dialog
   Caption = "Standard Cost"
   ClientHeight = 4230
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 7080
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 4230
   ScaleWidth = 7080
   ShowInTaskbar = 0 'False
   Begin VB.CheckBox optVew
      Height = 255
      Left = 360
      TabIndex = 29
      Top = 0
      Visible = 0 'False
      Width = 735
   End
   Begin VB.TextBox cmbPrt
      Height = 285
      Left = 1440
      TabIndex = 0
      Tag = "3"
      Top = 720
      Width = 3015
   End
   Begin VB.CommandButton cmdFnd
      Height = 315
      Left = 4560
      Picture = "diaSCe01a.frx":0000
      Style = 1 'Graphical
      TabIndex = 28
      TabStop = 0 'False
      ToolTipText = "Find A Part"
      Top = 720
      UseMaskColor = -1 'True
      Width = 350
   End
   Begin VB.CommandButton cmdUpd
      Caption = "&Update"
      Enabled = 0 'False
      Height = 315
      Left = 6120
      TabIndex = 27
      ToolTipText = "Update Standard Cost To Calculated Total"
      Top = 600
      Width = 875
   End
   Begin VB.TextBox txtCst
      Alignment = 1 'Right Justify
      Height = 285
      Left = 5880
      TabIndex = 6
      Tag = "1"
      Top = 3600
      Width = 1035
   End
   Begin VB.TextBox txtHrs
      Alignment = 1 'Right Justify
      Height = 285
      Left = 5880
      TabIndex = 1
      Tag = "1"
      Top = 1200
      Width = 1035
   End
   Begin VB.TextBox txtOhd
      Alignment = 1 'Right Justify
      Height = 285
      Left = 5880
      TabIndex = 5
      Tag = "1"
      Top = 2640
      Width = 1035
   End
   Begin VB.TextBox txtMat
      Alignment = 1 'Right Justify
      Height = 285
      Left = 5880
      TabIndex = 3
      Tag = "1"
      Top = 1920
      Width = 1035
   End
   Begin VB.TextBox txtExp
      Alignment = 1 'Right Justify
      Height = 285
      Left = 5880
      TabIndex = 4
      Tag = "1"
      Top = 2280
      Width = 1035
   End
   Begin VB.TextBox txtLab
      Alignment = 1 'Right Justify
      Height = 285
      Left = 5880
      TabIndex = 2
      Tag = "1"
      Top = 1560
      Width = 1035
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 6120
      TabIndex = 7
      TabStop = 0 'False
      Top = 120
      Width = 875
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 360
      Top = 2400
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 4230
      FormDesignWidth = 7080
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 26
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaSCe01a.frx":0342
      PictureDn = "diaSCe01a.frx":0488
   End
   Begin VB.Line Line1
      X1 = 4560
      X2 = 6960
      Y1 = 3000
      Y2 = 3000
   End
   Begin VB.Label lblRev
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1440
      TabIndex = 25
      Top = 3360
      Width = 855
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Last Revised"
      Height = 285
      Index = 13
      Left = 120
      TabIndex = 24
      Top = 3360
      Width = 1275
   End
   Begin VB.Label lblMbe
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 3480
      TabIndex = 23
      Top = 3000
      Width = 375
   End
   Begin VB.Label lblLvl
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1440
      TabIndex = 22
      Top = 3000
      Width = 375
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Make/Buy/Either"
      Height = 285
      Index = 11
      Left = 2160
      TabIndex = 21
      Top = 3000
      Width = 1275
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Type"
      Height = 285
      Index = 10
      Left = 120
      TabIndex = 20
      Top = 3000
      Width = 1275
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Current Standard Cost"
      Height = 405
      Index = 9
      Left = 4680
      TabIndex = 19
      Top = 3480
      Width = 1155
   End
   Begin VB.Label lblTot
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 5880
      TabIndex = 18
      Tag = "1"
      Top = 3120
      Width = 1035
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Total"
      Height = 285
      Index = 8
      Left = 4680
      TabIndex = 17
      Top = 3120
      Width = 1035
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Hours"
      Height = 285
      Index = 7
      Left = 4680
      TabIndex = 16
      Top = 1200
      Width = 915
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Overhead"
      Height = 285
      Index = 6
      Left = 4680
      TabIndex = 15
      Top = 2640
      Width = 915
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Material"
      Height = 285
      Index = 5
      Left = 4680
      TabIndex = 14
      Top = 1920
      Width = 1035
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Expense"
      Height = 285
      Index = 4
      Left = 4680
      TabIndex = 13
      Top = 2280
      Width = 1035
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Labor"
      Height = 285
      Index = 3
      Left = 4680
      TabIndex = 12
      Top = 1560
      Width = 1035
   End
   Begin VB.Label lblExt
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 1335
      Left = 1440
      TabIndex = 11
      Top = 1440
      Width = 3015
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1440
      TabIndex = 10
      Top = 1080
      Width = 3015
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Number"
      Height = 285
      Index = 0
      Left = 120
      TabIndex = 9
      Top = 720
      Width = 1275
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Description"
      Height = 285
      Index = 2
      Left = 120
      TabIndex = 8
      Top = 1080
      Width = 1035
   End
End
Attribute VB_Name = "diaSCe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'***********************************************************************************
' diaCost21 - Revise Part Standard Product Cost
'
' Created: (cjs)
' Revisions:
'   03/28/02 (nth) Removed the leading charature search from combo box
'   04/08/02 (nth) Changed update database columns from *LEV* to *TOT*
'   04/08/02 (nth) Changed the enabling and disabling of OH,HRS,LAB,EXP,MAT per JLH
'   07/21/02 (nth) Added Part lookup
'
'***********************************************************************************

Dim rdoPrt As rdoResultset
Dim rdoQry As rdoQuery
Dim bOnLoad As Byte
Dim bGoodPart As Byte
Dim bClose As Byte

Dim nCurStdCost As Single

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'***********************************************************************************

Private Sub cmbprt_LostFocus()
   cmbprt = CheckLen(cmbprt, 30)
   If Len(cmbprt) Then bGoodPart = GetPart()
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bClose = True
End Sub

Private Sub cmdFnd_Click()
   optVew.Value = vbChecked
   VewParts.Show
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Standard Cost"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdUpd_Click()
   If bGoodPart Then
      Err = 0
      On Error Resume Next
      rdoPrt.Edit
      rdoPrt!PASTDCOST = Val(lblTot)
      rdoPrt!PAREVDATE = Format(Now, "m/d/yy")
      rdoPrt.Update
      If Err > 0 Then
         ValidateEdit Me
      Else
         txtCst = lblTot
         Sysmsg "Standard Cost Updated.", True
         lblRev = Format(Now, "mm/dd/yy")
      End If
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      cmbprt = Cur.CurrentPart
      If Len(cmbprt) Then
         bGoodPart = GetPart
      End If
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAREVDATE," _
          & "PAMAKEBUY,PATOTLABOR,PATOTEXP,PATOTMATL,PATOTOH," _
          & "PATOTHRS,PASTDCOST,PAEXTDESC FROM PartTable WHERE PARTREF= ? "
   Set rdoQry = RdoCon.CreateQuery("", sSql)
   bOnLoad = True
   bClose = False
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bGoodPart = 1 Then
      Cur.CurrentPart = Trim(cmbprt)
      SaveCurrentSelections
   End If
   FormUnload
   Set rdoQry = Nothing
   Set rdoPrt = Nothing
   Set diaSCe01a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
   txtLab = "0.000"
   txtMat = "0.000"
   txtExp = "0.000"
   txtOhd = "0.000"
   txtHrs = "0.000"
   lblTot = "0.000"
   txtCst = "0.000"
End Sub

Public Function GetPart() As Byte
   Dim sPartRef As String
   sPartRef = Compress(cmbprt)
   rdoQry(0) = sPartRef
   On Error GoTo DiaErr1
   bSqlRows = GetQuerySet(rdoPrt, rdoQry, ES_KEYSET)
   
   If bSqlRows Then
      With rdoPrt
         cmbprt = "" & Trim(!PARTNUM)
         lblDsc = "" & Trim(!PADESC)
         lblExt = "" & Trim(!PAEXTDESC)
         lblLvl = Format(!PALEVEL, "0")
         lblMbe = "" & Trim(!PAMAKEBUY)
         lblRev = Format(!PAREVDATE, "mm/dd/yy")
         
         ' Clean out any nulls if we have them
         .Edit
         If IsNull(!PATOTLABOR) Then !PATOTLABOR = 0
         If IsNull(!PATOTEXP) Then !PATOTEXP = 0
         If IsNull(!PATOTMATL) Then !PATOTMATL = 0
         If IsNull(!PATOTOH) Then !PATOTOH = 0
         If IsNull(!PATOTHRS) Then !PATOTHRS = 0
         If IsNull(!PASTDCOST) Then !PASTDCOST = 0
         .Update
         
         txtLab = Format(!PATOTLABOR, "#,###,##0.000")
         txtExp = Format(!PATOTEXP, "#,###,##0.000")
         txtMat = Format(!PATOTMATL, "#,###,##0.000")
         txtOhd = Format(!PATOTOH, "#,###,##0.000")
         txtHrs = Format(!PATOTHRS, "#,###,##0.000")
         txtCst = Format(!PASTDCOST, "#,###,##0.000")
         
         nCurStdCost = Val(txtCst)
         .Cancel
      End With
      
      Select Case Val(lblLvl)
         Case 7
            
            txtExp.BackColor = cmbprt.BackColor
            txtExp.Enabled = True
            txtLab.BackColor = Me.BackColor
            txtLab.Enabled = False
            txtOhd.BackColor = Me.BackColor
            txtOhd.Enabled = False
            txtHrs.BackColor = Me.BackColor
            txtHrs.Enabled = False
            txtMat.BackColor = Me.BackColor
            txtMat.Enabled = False
         Case 4
            txtExp.BackColor = Me.BackColor
            txtExp.Enabled = False
            txtLab.BackColor = Me.BackColor
            txtLab.Enabled = False
            txtOhd.BackColor = Me.BackColor
            txtOhd.Enabled = False
            txtHrs.BackColor = Me.BackColor
            txtHrs.Enabled = False
            txtMat.BackColor = cmbprt.BackColor
            txtMat.Enabled = True
         Case 8
            ' hmm
            
         Case Else
            txtExp.BackColor = cmbprt.BackColor
            txtExp.Enabled = True
            txtLab.BackColor = cmbprt.BackColor
            txtLab.Enabled = True
            txtOhd.BackColor = cmbprt.BackColor
            txtOhd.Enabled = True
            txtHrs.BackColor = cmbprt.BackColor
            txtHrs.Enabled = True
            txtMat.BackColor = cmbprt.BackColor
            txtMat.Enabled = True
      End Select
      
      UpdateTotals
      GetPart = 1
      cmdUpd.Enabled = True
   Else
      lblLvl = ""
      lblMbe = ""
      lblRev = ""
      txtLab = "0.000"
      txtMat = "0.000"
      txtExp = "0.000"
      txtOhd = "0.000"
      txtHrs = "0.000"
      lblTot = "0.000"
      txtCst = "0.000"
      cmdUpd.Enabled = False
      GetPart = 0
      lblDsc = "*** No Current Part ***"
      lblExt = ""
   End If
   Exit Function
   
   DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc_Change()
   If lblDsc = "*** No Current Part ***" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub


Public Sub UpdateTotals()
   Dim nTotal As Single
   nTotal = Val(txtExp) + Val(txtHrs) + Val(txtLab) + Val(txtMat) + Val(txtOhd)
   lblTot = Format(nTotal, "#,###,##0.000")
End Sub

Private Sub optVew_Click()
   If optVew.Value = vbUnchecked Then
      ' Part search is closing refresh form
      cmbprt_LostFocus
   End If
End Sub

Private Sub txtCst_LostFocus()
   If bClose = 0 Then
      txtCst = CheckLen(txtCst, 13)
      txtCst = Format(Abs(Val(txtCst)), "#,###,##0.000")
      
      ' did the std cost change ?
      If Val(txtCst) <> nCurStdCost Then
         UpdateTotals
         If bGoodPart Then
            On Error Resume Next
            rdoPrt.Edit
            rdoPrt!PASTDCOST = Val(txtCst)
            rdoPrt!PAREVDATE = Format(Now, "mm/dd/yy")
            rdoPrt.Update
            
            If Err = 0 Then
               ValidateEdit Me
               Sysmsg "Standard Cost Updated.", True
               lblRev = Format(Now, "mm/dd/yy")
               nCurStdCost = Val(txtCst)
            Else
               MsgBox "Couldn't Update Standard Cost For The Part.", _
                  vbInformation, Caption
               txtCst = Format(nCurStdCost, "#,###,##0.000")
            End If
         End If
      End If
   End If
End Sub


Private Sub txtExp_LostFocus()
   txtExp = CheckLen(txtExp, 13)
   txtExp = Format(Abs(Val(txtExp)), "#,###,##0.000")
   UpdateTotals
   If bGoodPart Then
      On Error Resume Next
      rdoPrt.Edit
      rdoPrt!PATOTEXP = Val(txtExp)
      rdoPrt.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtHrs_LostFocus()
   txtHrs = CheckLen(txtHrs, 13)
   txtHrs = Format(Abs(Val(txtHrs)), "#,###,##0.000")
   UpdateTotals
   If bGoodPart Then
      On Error Resume Next
      rdoPrt.Edit
      rdoPrt!PATOTHRS = Val(txtHrs)
      rdoPrt.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtLab_LostFocus()
   txtLab = CheckLen(txtLab, 13)
   txtLab = Format(Abs(Val(txtLab)), "#,###,##0.000")
   UpdateTotals
   If bGoodPart Then
      On Error Resume Next
      rdoPrt.Edit
      rdoPrt!PATOTLABOR = Val(txtLab)
      rdoPrt.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtMat_LostFocus()
   txtMat = CheckLen(txtMat, 13)
   txtMat = Format(Abs(Val(txtMat)), "#,###,##0.000")
   UpdateTotals
   If bGoodPart Then
      On Error Resume Next
      rdoPrt.Edit
      rdoPrt!PATOTMATL = Val(txtMat)
      rdoPrt.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub


Private Sub txtOhd_LostFocus()
   txtOhd = CheckLen(txtOhd, 13)
   txtOhd = Format(Abs(Val(txtOhd)), "#,###,##0.000")
   UpdateTotals
   If bGoodPart Then
      On Error Resume Next
      rdoPrt.Edit
      rdoPrt!PATOTOH = Val(txtOhd)
      rdoPrt.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub
