VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaPcInfo
   BorderStyle = 3 'Fixed Dialog
   Caption = "Cost Information"
   ClientHeight = 6315
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 8265
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 6315
   ScaleWidth = 8265
   ShowInTaskbar = 0 'False
   Begin VB.TextBox txtOhd
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 5
      Left = 7080
      TabIndex = 62
      Tag = "8"
      Top = 4560
      Width = 1035
   End
   Begin VB.TextBox txtMat
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 5
      Left = 7080
      TabIndex = 61
      Tag = "8"
      Top = 3840
      Width = 1035
   End
   Begin VB.TextBox txtExp
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 5
      Left = 7080
      TabIndex = 60
      Tag = "8"
      Top = 4200
      Width = 1035
   End
   Begin VB.TextBox txtLab
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 5
      Left = 7080
      TabIndex = 59
      Tag = "8"
      Top = 3480
      Width = 1035
   End
   Begin VB.TextBox txthrs
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 5
      Left = 7080
      TabIndex = 58
      Tag = "8"
      Top = 3120
      Width = 1035
   End
   Begin VB.TextBox txtOhd
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 4
      Left = 6000
      TabIndex = 57
      Tag = "8"
      Top = 4560
      Width = 1035
   End
   Begin VB.TextBox txtMat
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 4
      Left = 6000
      TabIndex = 56
      Tag = "8"
      Top = 3840
      Width = 1035
   End
   Begin VB.TextBox txtExp
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 4
      Left = 6000
      TabIndex = 55
      Tag = "8"
      Top = 4200
      Width = 1035
   End
   Begin VB.TextBox txtLab
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 4
      Left = 6000
      TabIndex = 54
      Tag = "8"
      Top = 3480
      Width = 1035
   End
   Begin VB.TextBox txthrs
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 4
      Left = 6000
      TabIndex = 53
      Tag = "8"
      Top = 3120
      Width = 1035
   End
   Begin VB.CommandButton cmdStd
      Enabled = 0 'False
      Height = 315
      Left = 1680
      TabIndex = 50
      ToolTipText = "Convert Purposed To Std Cost"
      Top = 5880
      Width = 875
   End
   Begin VB.TextBox txthrs
      Alignment = 1 'Right Justify
      Height = 285
      Index = 3
      Left = 4920
      TabIndex = 45
      Tag = "1"
      Top = 3120
      Width = 1035
   End
   Begin VB.TextBox txthrs
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 2
      Left = 3840
      TabIndex = 44
      Tag = "8"
      Top = 3120
      Width = 1035
   End
   Begin VB.TextBox txthrs
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 1
      Left = 2760
      TabIndex = 43
      Tag = "1"
      Top = 3120
      Width = 1035
   End
   Begin VB.TextBox txthrs
      Alignment = 1 'Right Justify
      Enabled = 0 'False
      Height = 285
      Index = 0
      Left = 1680
      TabIndex = 42
      Tag = "1"
      Top = 3120
      Width = 1035
   End
   Begin VB.TextBox txtLab
      Alignment = 1 'Right Justify
      Enabled = 0 'False
      Height = 285
      Index = 0
      Left = 1680
      TabIndex = 2
      Tag = "1"
      Top = 3480
      Width = 1035
   End
   Begin VB.TextBox txtLab
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 1
      Left = 2760
      TabIndex = 6
      Tag = "1"
      Top = 3480
      Width = 1035
   End
   Begin VB.TextBox txtLab
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 2
      Left = 3840
      TabIndex = 10
      Tag = "8"
      Top = 3480
      Width = 1035
   End
   Begin VB.TextBox txtLab
      Alignment = 1 'Right Justify
      Height = 285
      Index = 3
      Left = 4920
      TabIndex = 14
      Tag = "1"
      Top = 3480
      Width = 1035
   End
   Begin VB.TextBox txtStd
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Left = 6000
      TabIndex = 19
      Tag = "8"
      Top = 5040
      Width = 1035
   End
   Begin VB.TextBox txtBud
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Left = 3840
      TabIndex = 18
      Tag = "8"
      Top = 5040
      Width = 1035
   End
   Begin VB.CommandButton cmdUpd
      Enabled = 0 'False
      Height = 315
      Left = 1680
      TabIndex = 20
      ToolTipText = "Update Standard Cost To Calculated Total"
      Top = 5520
      Width = 875
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 7680
      Top = 600
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 6315
      FormDesignWidth = 8265
   End
   Begin VB.TextBox txtExp
      Alignment = 1 'Right Justify
      Height = 285
      Index = 3
      Left = 4920
      TabIndex = 16
      Tag = "1"
      Top = 4200
      Width = 1035
   End
   Begin VB.TextBox txtMat
      Alignment = 1 'Right Justify
      Height = 285
      Index = 3
      Left = 4920
      TabIndex = 15
      Tag = "1"
      Top = 3840
      Width = 1035
   End
   Begin VB.TextBox txtOhd
      Alignment = 1 'Right Justify
      Height = 285
      Index = 3
      Left = 4920
      TabIndex = 17
      Tag = "1"
      Top = 4560
      Width = 1035
   End
   Begin VB.TextBox txtExp
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 2
      Left = 3840
      TabIndex = 12
      Tag = "8"
      Top = 4200
      Width = 1035
   End
   Begin VB.TextBox txtMat
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 2
      Left = 3840
      TabIndex = 11
      Tag = "8"
      Top = 3840
      Width = 1035
   End
   Begin VB.TextBox txtOhd
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 2
      Left = 3840
      TabIndex = 13
      Tag = "8"
      Top = 4560
      Width = 1035
   End
   Begin VB.TextBox txtExp
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 1
      Left = 2760
      TabIndex = 8
      Tag = "1"
      Top = 4200
      Width = 1035
   End
   Begin VB.TextBox txtMat
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 1
      Left = 2760
      TabIndex = 7
      Tag = "1"
      Top = 3840
      Width = 1035
   End
   Begin VB.TextBox txtOhd
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Index = 1
      Left = 2760
      TabIndex = 9
      Tag = "1"
      Top = 4560
      Width = 1035
   End
   Begin VB.TextBox txtExp
      Alignment = 1 'Right Justify
      Enabled = 0 'False
      Height = 285
      Index = 0
      Left = 1680
      TabIndex = 4
      Tag = "1"
      Top = 4200
      Width = 1035
   End
   Begin VB.TextBox txtMat
      Alignment = 1 'Right Justify
      Enabled = 0 'False
      Height = 285
      Index = 0
      Left = 1680
      TabIndex = 3
      Tag = "1"
      Top = 3840
      Width = 1035
   End
   Begin VB.TextBox txtOhd
      Alignment = 1 'Right Justify
      Enabled = 0 'False
      Height = 285
      Index = 0
      Left = 1680
      TabIndex = 5
      Tag = "1"
      Top = 4560
      Width = 1035
   End
   Begin VB.CommandButton cmdSel
      Caption = "S&elect"
      Height = 315
      Left = 4800
      TabIndex = 1
      ToolTipText = "Search For Part(s)"
      Top = 360
      Width = 855
   End
   Begin VB.ComboBox cmbPrt
      DataSource = "rDt1"
      Height = 315
      Left = 1440
      Sorted = -1 'True
      TabIndex = 0
      Tag = "3"
      ToolTipText = "Enter At Least (1) Leading Character"
      Top = 360
      Width = 3255
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 7320
      TabIndex = 21
      TabStop = 0 'False
      Top = 90
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 37
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
      PictureUp = "diaPcInfo.frx":0000
      PictureDn = "diaPcInfo.frx":0146
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Update Standard"
      Height = 285
      Index = 19
      Left = 120
      TabIndex = 52
      Top = 5880
      Width = 1755
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Update Purposed"
      Height = 285
      Index = 18
      Left = 120
      TabIndex = 51
      Top = 5520
      Width = 1635
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Total"
      Height = 285
      Index = 17
      Left = 120
      TabIndex = 49
      Top = 5040
      Width = 1035
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Previous"
      Height = 285
      Index = 14
      Left = 7200
      TabIndex = 48
      Top = 2760
      Width = 915
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Std Cost"
      Height = 285
      Index = 13
      Left = 6000
      TabIndex = 47
      Top = 2760
      Width = 1035
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Hours"
      Height = 285
      Index = 16
      Left = 120
      TabIndex = 46
      Top = 3120
      Width = 1275
   End
   Begin VB.Label lblExt
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 975
      Left = 1440
      TabIndex = 41
      Top = 1080
      Width = 5295
   End
   Begin VB.Label lblCnt
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Caption = "0"
      Height = 255
      Left = 6120
      TabIndex = 40
      ToolTipText = "Counted"
      Top = 720
      Width = 615
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Count"
      Height = 285
      Index = 15
      Left = 4800
      TabIndex = 39
      Top = 720
      Width = 1275
   End
   Begin VB.Line Line2
      X1 = 1560
      X2 = 1560
      Y1 = 2640
      Y2 = 5280
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Labor Cost"
      Height = 285
      Index = 1
      Left = 120
      TabIndex = 38
      Top = 3480
      Width = 1275
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Type"
      Height = 285
      Index = 12
      Left = 120
      TabIndex = 36
      Top = 2160
      Width = 1395
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Make/Buy/Either"
      Height = 285
      Index = 11
      Left = 2400
      TabIndex = 35
      Top = 2160
      Width = 1275
   End
   Begin VB.Label lblLvl
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1440
      TabIndex = 34
      Top = 2160
      Width = 375
   End
   Begin VB.Label lblMbe
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 4080
      TabIndex = 33
      Top = 2160
      Width = 375
   End
   Begin VB.Line Line1
      X1 = 120
      X2 = 8160
      Y1 = 3000
      Y2 = 3000
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Purposed"
      Height = 285
      Index = 10
      Left = 4920
      TabIndex = 32
      Top = 2760
      Width = 1035
   End
   Begin VB.Label z1
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      Caption = "Total"
      Height = 285
      Index = 9
      Left = 3840
      TabIndex = 31
      Top = 2760
      Width = 1035
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Lower Levels"
      Height = 285
      Index = 8
      Left = 2760
      TabIndex = 30
      Top = 2760
      Width = 1275
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "This Level"
      Height = 285
      Index = 7
      Left = 1680
      TabIndex = 29
      Top = 2760
      Width = 1035
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Proposed Cost Components"
      Height = 525
      Index = 2
      Left = 120
      TabIndex = 28
      Top = 2580
      Width = 1395
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Expense Cost"
      Height = 285
      Index = 0
      Left = 120
      TabIndex = 27
      Top = 4200
      Width = 1395
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Material Cost"
      Height = 285
      Index = 5
      Left = 120
      TabIndex = 26
      Top = 3840
      Width = 1275
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Overhead Cost"
      Height = 285
      Index = 6
      Left = 120
      TabIndex = 25
      Top = 4560
      Width = 1395
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Description"
      Height = 285
      Index = 4
      Left = 120
      TabIndex = 24
      Top = 720
      Width = 1395
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Number"
      Height = 285
      Index = 3
      Left = 120
      TabIndex = 23
      Top = 360
      Width = 1395
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1440
      TabIndex = 22
      Top = 720
      Width = 3015
   End
End
Attribute VB_Name = "diaPcInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
' diaPcInfo - Cost Information
'
' Created: 11/30/01 (nth)
' Revisions:
'   Roll up lower level cost from BOM 11/21/01 (nth)
'
'*********************************************************************************
'
Option Explicit

Dim RdoPrt As rdoResultset
Dim RdoQry As rdoQuery
Dim bOnLoad As Byte
Dim bGoodPart As Byte

' New method of preserving colors - giving it a try
Dim lForeColor As Long
Dim lBackColor As Long

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   If cmbPrt.ListCount > 0 Then bGoodPart = GetPart()
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If cmbPrt.ListCount Then bGoodPart = GetPart()
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdSel_Click()
   FillCombo
End Sub

Private Sub cmdStd_Click()
   UpdateStd
End Sub

Private Sub cmdUpd_Click()
   UpdatePurposed
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
      
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   
   lForeColor = Me.ForeColor
   lBackColor = Me.BackColor
   
   sSql = "SELECT PARTREF, PARTNUM, PADESC, PALEVEL, PAREVDATE," _
          & "PAEXTDESC, PAMAKEBUY, PALEVLABOR, PALEVEXP, PALEVMATL, PALEVOH," _
          & "PALEVHRS, PASTDCOST, PABOMLABOR, PABOMEXP, PABOMMATL, PABOMOH," _
          & "PABOMHRS, PABOMREV, PAPREVLABOR, PAPREVEXP, PAPREVMATL, PAPREVOH," _
          & "PAPREVHRS, PATOTHRS, PATOTEXP, PATOTLABOR, PATOTMATL, PATOTOH,PAROUTING " _
          & "FROM PartTable WHERE PARTREF = ?"
   Set RdoQry = RdoCon.CreateQuery("", sSql)
   bOnLoad = True
   
   txthrs(0).BackColor = Me.BackColor
   txtOhd(0).BackColor = Me.BackColor
   txtLab(0).BackColor = Me.BackColor
   txtExp(0).BackColor = Me.BackColor
   txtMat(0).BackColor = Me.BackColor
   
   txthrs(1).BackColor = Me.BackColor
   txtOhd(1).BackColor = Me.BackColor
   txtLab(1).BackColor = Me.BackColor
   txtExp(1).BackColor = Me.BackColor
   txtMat(1).BackColor = Me.BackColor
   
   txthrs(2).BackColor = Me.BackColor
   txtOhd(2).BackColor = Me.BackColor
   txtLab(2).BackColor = Me.BackColor
   txtExp(2).BackColor = Me.BackColor
   txtMat(2).BackColor = Me.BackColor
   
   txthrs(4).BackColor = Me.BackColor
   txtOhd(4).BackColor = Me.BackColor
   txtLab(4).BackColor = Me.BackColor
   txtExp(4).BackColor = Me.BackColor
   txtMat(4).BackColor = Me.BackColor
   
   txthrs(5).BackColor = Me.BackColor
   txtOhd(5).BackColor = Me.BackColor
   txtLab(5).BackColor = Me.BackColor
   txtExp(5).BackColor = Me.BackColor
   txtMat(5).BackColor = Me.BackColor
   
   txtBud.BackColor = Me.BackColor
   txtStd.BackColor = Me.BackColor
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaPcInfo = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub txtExp_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtExp_LostFocus(Index As Integer)
   txtExp(Index) = CheckLen(txtExp(Index), 11)
   txtExp(Index) = Format(Abs(Val(txtExp(Index))), "#,###,##0.000")
   UpdateTotals
End Sub

Private Sub txthrs_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtHrs_LostFocus(Index As Integer)
   txthrs(Index) = CheckLen(txthrs(Index), 11)
   txthrs(Index) = Format(Abs(Val(txthrs(Index))), "#,###,##0.000")
   UpdateTotals
End Sub

Private Sub txtLab_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtLab_LostFocus(Index As Integer)
   txtLab(Index) = CheckLen(txtLab(Index), 11)
   txtLab(Index) = Format(Abs(Val(txtLab(Index))), "#,###,##0.000")
   UpdateTotals
End Sub

Private Sub txtMat_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtMat_LostFocus(Index As Integer)
   txtMat(Index) = CheckLen(txtMat(Index), 11)
   txtMat(Index) = Format(Abs(Val(txtMat(Index))), "#,###,##0.000")
   UpdateTotals
End Sub

Private Sub txtOhd_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Public Sub FillCombo()
   Dim i As Integer
   Dim rdocmb As rdoResultset
   Dim sPartRef As String
   On Error GoTo DiaErr1
   
   sPartRef = Compress(cmbPrt)
   cmbPrt.Clear
   If Len(sPartRef) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE " _
             & "PARTREF LIKE '" & sPartRef & "%'"
      bSqlRows = GetDataSet(rdocmb, ES_FORWARD)
      If bSqlRows Then
         With rdocmb
            Do Until .EOF
               i = i + 1
               If i > 299 Then Exit Do
               cmbPrt.AddItem "" & Trim(!PARTNUM)
               .MoveNext
            Loop
            .Cancel
         End With
         lblCnt = i
         If cmbPrt.ListCount > 0 Then
            cmbPrt = cmbPrt.List(0)
            bGoodPart = GetPart()
         End If
      End If
   Else
      lblCnt = 0
      MsgBox "Please Select (3) Or More Leading Characters.", _
         vbInformation, Caption
   End If
   Set rdocmb = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Function GetPart() As Byte
   Dim sPartRef As String
   Dim bHasBOM As Byte
   Dim rdoBOM As rdoResultset
   Dim rdoLower As rdoResultset
   Dim cPABOMLABOR As Currency
   Dim cPABOMEXP As Currency
   Dim cPABOMMATL As Currency
   Dim cPABOMOH As Currency
   Dim cPABOMHRS As Currency
   
   sPartRef = Compress(cmbPrt)
   
   On Error GoTo DiaErr1
   
   ' Check for a BOM
   sSql = "SELECT BMPARTREF FROM BmplTable WHERE BMASSYPART = '" & sPartRef & "'"
   bHasBOM = GetDataSet(rdoBOM)
   If bHasBOM Then
      With rdoBOM
         While Not .EOF
            ' Roll up lower level cost
            RdoQry(0) = Trim(!BMPARTREF)
            bSqlRows = GetQuerySet(rdoLower, RdoQry, ES_FORWARD)
            cPABOMLABOR = cPABOMLABOR + rdoLower!PALEVLABOR
            cPABOMEXP = cPABOMEXP + rdoLower!PALEVEXP
            cPABOMMATL = cPABOMMATL + rdoLower!PALEVMATL
            cPABOMOH = cPABOMOH + rdoLower!PALEVOH
            cPABOMHRS = cPABOMHRS + rdoLower!PALEVHRS
            .MoveNext
            Set rdoLower = Nothing
         Wend
      End With
      Set rdoBOM = Nothing
      Set rdoLower = Nothing
   End If
   
   RdoQry(0) = sPartRef
   bSqlRows = GetQuerySet(RdoPrt, RdoQry, ES_KEYSET)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PARTNUM)
         lblDsc = lForeColor
         lblDsc = "" & Trim(!PADESC)
         lblExt = "" & Trim(!PAEXTDESC)
         lblLvl = Format(!PALEVEL, "0")
         lblMbe = "" & Trim(!PAMAKEBUY)
         
         ' This Level
         txtLab(0) = Format(Val("" & !PALEVLABOR), "#,###,##0.000")
         txtExp(0) = Format(Val("" & !PALEVEXP), "#,###,##0.000")
         txtMat(0) = Format(Val("" & !PALEVMATL), "#,###,##0.000")
         txtOhd(0) = Format(Val("" & !PALEVOH), "#,###,##0.000")
         txthrs(0) = Format(Val("" & !PALEVHRS), "#,###,##0.000")
         
         ' Lower Levels
         txtLab(1) = Format(Val(cPABOMLABOR), "#,###,##0.000")
         txtExp(1) = Format(Val(cPABOMEXP), "#,###,##0.000")
         txtMat(1) = Format(Val(cPABOMMATL), "#,###,##0.000")
         txtOhd(1) = Format(Val(cPABOMOH), "#,###,##0.000")
         txthrs(1) = Format(Val(cPABOMHRS), "#,###,##0.000")
         
         ' Total
         UpdateTotals
         
         ' Purposed
         txtLab(3) = txtLab(2)
         txtExp(3) = txtExp(2)
         txtMat(3) = txtMat(2)
         txtOhd(3) = txtOhd(2)
         txthrs(3) = txthrs(2)
         
         ' Std
         txtLab(4) = Format(Val("" & !PATOTLABOR), "#,###,##0.000")
         txtExp(4) = Format(Val("" & !PATOTEXP), "#,###,##0.000")
         txtMat(4) = Format(Val("" & !PATOTMATL), "#,###,##0.000")
         txtOhd(4) = Format(Val("" & !PATOTOH), "#,###,##0.000")
         txthrs(4) = Format(Val("" & !PATOTHRS), "#,###,##0.000")
         
         txtStd = Format(Val(txtLab(4)) + Val(txtExp(4)) + Val(txthrs(4)) _
                  + Val(txtMat(4)) + Val(txtOhd(4)), "#,###,##0.000")
         
         ' Previous
         txtLab(5) = Format(Val("" & !PAPREVLABOR), "#,###,##0.000")
         txtExp(5) = Format(Val("" & !PAPREVEXP), "#,###,##0.000")
         txtMat(5) = Format(Val("" & !PAPREVMATL), "#,###,##0.000")
         txtOhd(5) = Format(Val("" & !PAPREVOH), "#,###,##0.000")
         txthrs(5) = Format(Val("" & !PAPREVHRS), "#,###,##0.000")
         
         .Cancel
         
         Select Case lblLvl
            Case "4"
               If IsNull(!PAROUTING) And Not bHasBOM Then
                  txtExp(0).BackColor = Me.BackColor
                  txtLab(0).BackColor = Me.BackColor
                  txtOhd(0).BackColor = Me.BackColor
                  txthrs(0).BackColor = Me.BackColor
                  txtMat(0).BackColor = cmbPrt.BackColor
                  txtMat(0).Enabled = True
               End If
            Case Else
               If IsNull(!PAROUTING) Then
                  txtExp(0).BackColor = cmbPrt.BackColor
                  txtExp(0).Enabled = True
                  txtLab(0).BackColor = cmbPrt.BackColor
                  txtLab(0).Enabled = True
                  txtOhd(0).BackColor = cmbPrt.BackColor
                  txtOhd(0).Enabled = True
                  txthrs(0).BackColor = cmbPrt.BackColor
                  txthrs(0).Enabled = True
                  txtMat(0).BackColor = cmbPrt.BackColor
                  txtMat(0).Enabled = True
                  txtMat(0).BackColor = Me.BackColor
                  txtMat(0).Enabled = False
               End If
         End Select
      End With
      
      UpdateTotals
      GetPart = 1
      cmdUpd.Enabled = True
      cmdStd.Enabled = True
   Else
      cmdUpd.Enabled = False
      cmdStd.Enabled = False
      GetPart = 0
      lblDsc.ForeColor = ES_RED
      lblDsc = "*** No Current Part ***"
   End If
   Exit Function
   DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub UpdateTotals()
   txtLab(2) = Format(Val(txtLab(0)) + Val(txtLab(1)), "#,###,##0.000")
   txtExp(2) = Format(Val(txtExp(0)) + Val(txtExp(1)), "#,###,##0.000")
   txtMat(2) = Format(Val(txtMat(0)) + Val(txtMat(1)), "#,###,##0.000")
   txtOhd(2) = Format(Val(txtOhd(0)) + Val(txtOhd(1)), "#,###,##0.000")
   txthrs(2) = Format(Val(txthrs(0)) + Val(txthrs(1)), "#,###,##0.000")
   txtBud = Format(Val(txtLab(2)) + Val(txtExp(2)) + Val(txthrs(2)) _
            + Val(txtMat(2)) + Val(txtOhd(2)), "#,###,##0.000")
End Sub

Private Sub txtOhd_LostFocus(Index As Integer)
   txtOhd(Index) = CheckLen(txtOhd(Index), 11)
   txtOhd(Index) = Format(Abs(Val(txtOhd(Index))), "#,###,##0.000")
   UpdateTotals
End Sub

Private Sub UpdatePurposed()
   On Error Resume Next
   
   Err = 0
   With RdoPrt
      .Edit
      !PALEVHRS = Val(txthrs(0))
      !PALEVEXP = Val(txtExp(0))
      !PALEVMATL = Val(txtMat(0))
      !PALEVOH = Val(txtOhd(0))
      !PALEVLABOR = Val(txtLab(0))
      
      !PABOMHRS = Val(txthrs(1))
      !PABOMEXP = Val(txtExp(1))
      !PABOMMATL = Val(txtMat(1))
      !PABOMOH = Val(txtOhd(1))
      !PABOMLABOR = Val(txtLab(1))
      .Update
   End With
   
   If Err > 0 Then
      ValidateEdit Me
   Else
      On Error GoTo 0
      ' Calc the new purposed cost
      txtLab(3) = Format(Val(txtLab(0)) + Val(txtLab(1)), "#,###,##0.000")
      txtExp(3) = Format(Val(txtExp(0)) + Val(txtExp(1)), "#,###,##0.000")
      txtMat(3) = Format(Val(txtMat(0)) + Val(txtMat(1)), "#,###,##0.000")
      txtOhd(3) = Format(Val(txtOhd(0)) + Val(txtOhd(1)), "#,###,##0.000")
      txthrs(3) = Format(Val(txthrs(0)) + Val(txthrs(1)), "#,###,##0.000")
      Sysmsg "Purposed Cost Updated.", True
   End If
End Sub

Private Sub UpdateStd()
   Dim iResponse As Integer
   iResponse = MsgBox("Update Purposed To Standard Cost?", ES_YESQUESTION, Caption)
   If iResponse = vbNo Then
      Exit Sub
   End If
   
   On Error Resume Next
   Err = 0
   With RdoPrt
      .Edit
      ' Old Std cost
      !PAPREVHRS = Val(txthrs(4))
      !PAPREVEXP = Val(txtExp(4))
      !PAPREVMATL = Val(txtMat(4))
      !PAPREVOH = Val(txtOhd(4))
      !PAPREVLABOR = Val(txtLab(4))
      
      ' New Std cost from purposed
      !PATOTHRS = Val(txthrs(3))
      !PATOTEXP = Val(txtExp(3))
      !PATOTMATL = Val(txtMat(3))
      !PATOTOH = Val(txtOhd(3))
      !PATOTLABOR = Val(txtLab(3))
      .Update
   End With
   
   If Err > 0 Then
      ValidateEdit Me
   Else
      txtLab(4) = txtLab(3)
      txtExp(4) = txtExp(3)
      txtMat(4) = txtMat(3)
      txtOhd(4) = txtOhd(3)
      txthrs(4) = txthrs(3)
      txtStd = Format(Val(txtLab(4)) + Val(txtExp(4)) + Val(txthrs(4)) _
               + Val(txtMat(4)) + Val(txtOhd(4)), "#,###,##0.000")
      Sysmsg "Standard Cost Updated.", True
   End If
End Sub
