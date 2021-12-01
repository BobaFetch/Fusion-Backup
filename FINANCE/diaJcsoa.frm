VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form diaJcsoa
   BorderStyle = 3 'Fixed Dialog
   Caption = "Sales Order Allocations"
   ClientHeight = 5580
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 7425
   ControlBox = 0 'False
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 5580
   ScaleWidth = 7425
   ShowInTaskbar = 0 'False
   Begin VB.ComboBox cmbRun
      ForeColor = &H00800000&
      Height = 315
      Left = 5280
      TabIndex = 1
      Tag = "8"
      ToolTipText = "Select Run Number (Not CL or CA)"
      Top = 840
      Width = 1095
   End
   Begin VB.ComboBox cmbPrt
      Height = 315
      Left = 1200
      TabIndex = 0
      Tag = "3"
      ToolTipText = "Select MO Part Number (Not CA or CL)"
      Top = 480
      Width = 3545
   End
   Begin ComctlLib.ProgressBar prg1
      Height = 255
      Left = 1080
      TabIndex = 39
      Top = 5160
      Visible = 0 'False
      Width = 3915
      _ExtentX = 6906
      _ExtentY = 450
      _Version = 327682
      Appearance = 1
   End
   Begin VB.ComboBox cmbSon
      Height = 315
      Index = 4
      Left = 120
      TabIndex = 14
      Top = 4440
      Width = 915
   End
   Begin VB.TextBox txtQty
      Height = 315
      Index = 4
      Left = 5040
      TabIndex = 16
      Top = 4800
      Width = 915
   End
   Begin VB.ComboBox cmbItm
      Height = 315
      Index = 4
      Left = 1080
      TabIndex = 15
      Top = 4440
      Width = 735
   End
   Begin VB.ComboBox cmbSon
      Height = 315
      Index = 3
      Left = 120
      TabIndex = 11
      Top = 3720
      Width = 915
   End
   Begin VB.TextBox txtQty
      Height = 315
      Index = 3
      Left = 5040
      TabIndex = 13
      Top = 4080
      Width = 915
   End
   Begin VB.ComboBox cmbItm
      Height = 315
      Index = 3
      Left = 1080
      TabIndex = 12
      Top = 3720
      Width = 735
   End
   Begin VB.ComboBox cmbSon
      Height = 315
      Index = 2
      Left = 120
      TabIndex = 8
      Top = 3000
      Width = 915
   End
   Begin VB.TextBox txtQty
      Height = 315
      Index = 2
      Left = 5040
      TabIndex = 10
      Top = 3360
      Width = 915
   End
   Begin VB.ComboBox cmbItm
      Height = 315
      Index = 2
      Left = 1080
      TabIndex = 9
      Top = 3000
      Width = 735
   End
   Begin VB.ComboBox cmbSon
      Height = 315
      Index = 1
      Left = 120
      TabIndex = 5
      Top = 2280
      Width = 915
   End
   Begin VB.TextBox txtQty
      Height = 315
      Index = 1
      Left = 5040
      TabIndex = 7
      Top = 2640
      Width = 915
   End
   Begin VB.ComboBox cmbItm
      Height = 315
      Index = 1
      Left = 1080
      TabIndex = 6
      Top = 2280
      Width = 735
   End
   Begin VB.CommandButton cmdAll
      Caption = "Allocate"
      Height = 315
      Index = 0
      Left = 6480
      TabIndex = 18
      ToolTipText = "Allocate Selections"
      Top = 600
      Width = 875
   End
   Begin VB.ComboBox cmbItm
      ForeColor = &H00800000&
      Height = 315
      Index = 0
      Left = 1080
      TabIndex = 3
      ToolTipText = "Select Item From List"
      Top = 1560
      Width = 735
   End
   Begin VB.TextBox txtQty
      Height = 315
      Index = 0
      Left = 5040
      TabIndex = 4
      Top = 1920
      Width = 915
   End
   Begin VB.ComboBox cmbSon
      Height = 315
      Index = 0
      Left = 120
      TabIndex = 2
      ToolTipText = "Enter Or Select Sales Order "
      Top = 1560
      Width = 915
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 6480
      TabIndex = 17
      TabStop = 0 'False
      Top = 120
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 19
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
      PictureUp = "diaJcsoa.frx":0000
      PictureDn = "diaJcsoa.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 6600
      Top = 5040
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 5580
      FormDesignWidth = 7425
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1200
      TabIndex = 56
      Top = 840
      Width = 3255
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Run"
      Height = 255
      Index = 5
      Left = 4560
      TabIndex = 55
      Top = 840
      Width = 735
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Part Number"
      Height = 255
      Index = 4
      Left = 120
      TabIndex = 54
      Top = 480
      Width = 1095
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Selected     "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 7
      Left = 6120
      TabIndex = 53
      Top = 1920
      Visible = 0 'False
      Width = 855
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Run Qty      "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 6
      Left = 6120
      TabIndex = 52
      Top = 1320
      Width = 855
   End
   Begin VB.Label lblAllo
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 6120
      TabIndex = 51
      Top = 2160
      Visible = 0 'False
      Width = 855
   End
   Begin VB.Label lblRqty
      Alignment = 1 'Right Justify
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 6120
      TabIndex = 50
      Top = 1560
      Width = 855
   End
   Begin VB.Label lblRev
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Index = 4
      Left = 1080
      TabIndex = 49
      Top = 4800
      Width = 615
   End
   Begin VB.Label lblItm
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Index = 4
      Left = 120
      TabIndex = 48
      Top = 4800
      Width = 855
   End
   Begin VB.Label lblRev
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Index = 3
      Left = 1080
      TabIndex = 47
      Top = 4080
      Width = 615
   End
   Begin VB.Label lblItm
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Index = 3
      Left = 120
      TabIndex = 46
      Top = 4080
      Width = 855
   End
   Begin VB.Label lblRev
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Index = 2
      Left = 1080
      TabIndex = 45
      Top = 3360
      Width = 615
   End
   Begin VB.Label lblItm
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Index = 2
      Left = 120
      TabIndex = 44
      Top = 3360
      Width = 855
   End
   Begin VB.Label lblRev
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Index = 1
      Left = 1080
      TabIndex = 43
      Top = 2640
      Width = 615
   End
   Begin VB.Label lblItm
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Index = 1
      Left = 120
      TabIndex = 42
      Top = 2640
      Width = 855
   End
   Begin VB.Label lblRev
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Index = 0
      Left = 1080
      TabIndex = 41
      Top = 1920
      Width = 615
   End
   Begin VB.Label lblItm
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 40
      Top = 1920
      Width = 855
   End
   Begin VB.Label LblPrt
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 4
      Left = 1800
      TabIndex = 38
      Top = 4440
      Width = 3195
   End
   Begin VB.Label lblPDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 4
      Left = 1800
      TabIndex = 37
      Top = 4800
      Width = 3195
   End
   Begin VB.Label lblQty
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 4
      Left = 5040
      TabIndex = 36
      Top = 4440
      Width = 915
   End
   Begin VB.Label LblPrt
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 3
      Left = 1800
      TabIndex = 35
      Top = 3720
      Width = 3195
   End
   Begin VB.Label lblPDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 3
      Left = 1800
      TabIndex = 34
      Top = 4080
      Width = 3195
   End
   Begin VB.Label lblQty
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 3
      Left = 5040
      TabIndex = 33
      Top = 3720
      Width = 915
   End
   Begin VB.Label LblPrt
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 2
      Left = 1800
      TabIndex = 32
      Top = 3000
      Width = 3195
   End
   Begin VB.Label lblPDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 2
      Left = 1800
      TabIndex = 31
      Top = 3360
      Width = 3195
   End
   Begin VB.Label lblQty
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 2
      Left = 5040
      TabIndex = 30
      Top = 3000
      Width = 915
   End
   Begin VB.Label LblPrt
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 1
      Left = 1800
      TabIndex = 29
      Top = 2280
      Width = 3195
   End
   Begin VB.Label lblPDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 1
      Left = 1800
      TabIndex = 28
      Top = 2640
      Width = 3195
   End
   Begin VB.Label lblQty
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 1
      Left = 5040
      TabIndex = 27
      Top = 2280
      Width = 915
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Quantity       "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 3
      Left = 5040
      TabIndex = 26
      Top = 1320
      Width = 975
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Part Number/Description                                "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 2
      Left = 1800
      TabIndex = 25
      Top = 1320
      Width = 3255
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Item         "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 1
      Left = 1080
      TabIndex = 24
      Top = 1320
      Width = 735
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Sales Order "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 23
      Top = 1320
      Width = 975
   End
   Begin VB.Label lblQty
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 0
      Left = 5040
      TabIndex = 22
      Top = 1560
      Width = 915
   End
   Begin VB.Label lblPDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 0
      Left = 1800
      TabIndex = 21
      Top = 1920
      Width = 3195
   End
   Begin VB.Label LblPrt
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Index = 0
      Left = 1800
      TabIndex = 20
      Top = 1560
      Width = 3195
   End
End
Attribute VB_Name = "diaJcsoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

Dim RdoQry As rdoQuery
Dim RdoQry1 As rdoQuery
Dim RdoQry2 As rdoQuery
Dim bDataChanged As Byte
Dim bGoodItem As Byte
Dim bGoodList As Byte
Dim bOnLoad As Byte
Dim iIndex As Integer

Public Sub FillFormRuns()
   Dim RdoRns As rdoResultset
   Dim SPartRef As String
   ClearBoxes
   cmbRun.Clear
   SPartRef = Compress(cmbPrt)
   RdoQry(0) = SPartRef
   bSqlRows = GetQuerySet(RdoRns, RdoQry)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!RunNo, "####0")
            .MoveNext
         Loop
      End With
   End If
   If cmbRun.ListCount > 0 Then
      cmbRun = Format(cmbRun.List(0), "####0")
      GetCurrentAllocations
   Else
      ClearBoxes
   End If
   On Error Resume Next
   Set RdoRns = Nothing
   
   Exit Sub
   
   DiaErr1:
   sProcName = "fillformruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub cmbItm_Change(Index As Integer)
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub cmbItm_Click(Index As Integer)
   '    Dim sSoItem As String
   '    sSoItem = Trim(cmbItm(Index))
   '    lblItm(Index) = Val(sSoItem)
   '      If Len(sSoItem) > 0 Then
   '        If Asc(Right(sSoItem, 1)) > 64 Then
   '            lblRev(Index) = Right(sSoItem, 1)
   '        Else
   '            lblRev(Index) = ""
   '        End If
   '      End If
   bGoodItem = GetThisItem()
   
End Sub

Private Sub cmbItm_GotFocus(Index As Integer)
   txtQty(Index).Enabled = False
   iIndex = Index
   
End Sub


Private Sub cmbItm_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub cmbItm_LostFocus(Index As Integer)
   Dim sSoItem As String
   sSoItem = Trim(cmbItm(Index))
   lblItm(Index) = Val(sSoItem)
   If Len(sSoItem) > 0 Then
      If Asc(Right(sSoItem, 1)) > 64 Then
         lblRev(Index) = Right(sSoItem, 1)
      Else
         lblRev(Index) = ""
      End If
   End If
   bGoodItem = GetThisItem()
   On Error Resume Next
   If bGoodItem Then txtQty(Index).SetFocus
   
End Sub


Private Sub cmbPrt_Click()
   FindPart Me
   FillFormRuns
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   FindPart Me
   FillFormRuns
   
End Sub


Private Sub cmbRun_Click()
   If Val(cmbRun) > 0 Then GetCurrentAllocations _
          Else ClearBoxes
   
End Sub


Private Sub cmbRun_LostFocus()
   If Val(cmbRun) > 0 Then GetCurrentAllocations _
          Else ClearBoxes
   
End Sub


Private Sub cmbSon_Change(Index As Integer)
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub cmbSon_Click(Index As Integer)
   iIndex = Index
   bGoodList = GetSoItems()
   
End Sub

Private Sub cmbSon_GotFocus(Index As Integer)
   SelectFormat Me
   iIndex = Index
   If Len(Trim(cmbSon(Index))) = 0 Then txtQty(Index).Enabled = False
   
End Sub

Private Sub cmbSon_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub cmbSon_LostFocus(Index As Integer)
   cmbSon(Index) = CheckLen(cmbSon(Index), 6)
   bGoodList = GetSoItems()
   
End Sub

Private Sub CmdAll_Click(Index As Integer)
   If Val(lblAllo) > Val(lblRqty) Then
      Beep
      MsgBox "The Quantity Allocated Is Greater Than " _
         & vbCrLf & "The MO Quantity.  You Must Adjust.", vbExclamation, Caption
   Else
      AllocateItems
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      MouseCursor 13
      Fillcombo
      FillSalesOrders
      bOnLoad = False
      bDataChanged = False
   End If
   
End Sub

Private Sub Form_Load()
   Dim i As Integer
   SetDiaPos Me
   For i = 0 To 3
      lblItm(i).Visible = False
      lblRev(i).Visible = False
      txtQty(i) = Format(0)
      txtQty(i).Enabled = False
   Next
   lblItm(i).Visible = False
   lblRev(i).Visible = False
   txtQty(i) = Format(0)
   txtQty(i).Enabled = False
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
          & "RunsTable WHERE RUNREF = ? " _
          & "AND (RUNSTATUS<>'CA' AND RUNSTATUS<>'CL')  "
   Set RdoQry = RdoCon.CreateQuery("", sSql)
   
   sSql = "SELECT DISTINCT SONUMBER,ITSO,ITNUMBER,ITREV,ITACTUAL FROM " _
          & "SohdTable,SoitTable WHERE SONUMBER=ITSO AND (SONUMBER= ?" _
          & " AND ITACTUAL IS NULL)"
   Set RdoQry1 = RdoCon.CreateQuery("", sSql)
   
   sSql = "SELECT ITSO,ITNUMBER,ITREV,ITPART,ITQTY,PARTREF," _
          & "PARTNUM,PADESC FROM SoitTable,PartTable WHERE " _
          & "ITPART=PARTREF AND (ITSO= ? AND ITNUMBER=  ? AND ITREV= ? )"
   Set RdoQry2 = RdoCon.CreateQuery("", sSql)
   bOnLoad = True
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'diaSrvmo.optAll.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   rdoRes.Close
   Set RdoQry1 = Nothing
   Set RdoQry2 = Nothing
   FormUnload
   Set diaJcsoa = Nothing
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
End Sub

Private Sub txtQty_Change(Index As Integer)
   If Not bOnLoad Then bDataChanged = True
End Sub

Private Sub txtQty_GotFocus(Index As Integer)
   SelectFormat Me
   iIndex = Index
End Sub

Public Sub FillSalesOrders()
   Dim RdoSon As rdoResultset
   Dim a As Integer
   Dim i As Integer
   Dim iLastSo As Integer
   
   On Error GoTo DiaErr1
   a = 10
   sSql = "SELECT DISTINCT SONUMBER,SOTYPE,ITSO,ITNUMBER,ITACTUAL FROM " _
          & "SohdTable,SoitTable WHERE SONUMBER=ITSO AND ITACTUAL IS NULL"
   bSqlRows = GetDataSet(RdoSon)
   If bSqlRows Then
      iLastSo = -99
      With RdoSon
         Do Until .EOF
            If iLastSo <> !SONUMBER Then
               a = a + 5
               If a > 95 Then a = 95
               prg1.Value = a
               For i = 0 To 4
                  AddComboStr cmbSon(i).hWnd, "" & Trim(!SOTYPE) & Format$(!SONUMBER, "00000")
               Next
            End If
            iLastSo = !SONUMBER
            .MoveNext
         Loop
      End With
   End If
   prg1.Value = 100
   Set RdoSon = Nothing
   prg1.Visible = False
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   sProcName = "fillsalesor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function GetSoItems() As Byte
   Dim RdoItm As rdoResultset
   Dim lSonumber As Long
   Dim sSoItem As String
   
   lSonumber = Val(Right(cmbSon(iIndex), 5))
   On Error GoTo DiaErr1
   cmbItm(iIndex).Clear
   RdoQry1(0) = lSonumber
   bSqlRows = GetQuerySet(RdoItm, RdoQry1, ES_KEYSET)
   If bSqlRows Then
      With RdoItm
         cmbItm(iIndex) = "" & Str(!ITNUMBER) & Trim(!ITREV)
         Do Until .EOF
            'cmbItm(iIndex).AddItem "" & Str(!ITNUMBER) & Trim(!ITREV)
            AddComboStr cmbItm(iIndex).hWnd, "" & Format$(!ITNUMBER) & Trim(!ITREV)
            .MoveNext
         Loop
         .Cancel
      End With
      sSoItem = Trim(cmbItm(iIndex))
      lblItm(iIndex) = Val(sSoItem)
      If Len(sSoItem) > 0 Then
         If Asc(Right(sSoItem, 1)) > 64 Then
            lblRev(iIndex) = Right(sSoItem, 1)
         Else
            lblRev(iIndex) = ""
         End If
      End If
      txtQty(iIndex).Enabled = True
      GetSoItems = True
      bGoodItem = GetThisItem
   Else
      GetSoItems = False
   End If
   Set RdoItm = Nothing
   Exit Function
   
   DiaErr1:
   sProcName = "getsoitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Function GetThisItem()
   Dim RdoItm As rdoResultset
   Dim lSonumber As Long
   
   On Error GoTo DiaErr1
   lSonumber = Val(Right(cmbSon(iIndex), 5))
   RdoQry2(0) = lSonumber
   RdoQry2(1) = Val(lblItm(iIndex))
   RdoQry2(2) = "" & Trim(lblRev(iIndex))
   bSqlRows = GetQuerySet(RdoItm, RdoQry2)
   If bSqlRows Then
      With RdoItm
         LblPrt(iIndex) = "" & !PARTNUM
         lblPDsc(iIndex) = "" & !PADESC
         lblQty(iIndex) = Format(0 + !ITQTY, "####0")
         .Cancel
      End With
      txtQty(iIndex).Enabled = True
      GetThisItem = True
   Else
      txtQty(iIndex).Enabled = False
      GetThisItem = False
   End If
   Set RdoItm = Nothing
   Exit Function
   
   DiaErr1:
   sProcName = "getthisit"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtQty_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtQty_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtQty_LostFocus(Index As Integer)
   txtQty(Index) = CheckLen(txtQty(Index), 6)
   If Val(txtQty(Index)) > Val(lblQty(Index)) Then
      Beep
      txtQty(Index) = lblQty(Index)
   End If
   txtQty(Index) = Format(Abs(Val(txtQty(Index))), "####0")
   GetSelected
   
End Sub



Public Sub GetSelected()
   Dim i As Integer
   Dim l As Long
   On Error Resume Next
   For i = 0 To 3
      l = l + Val(txtQty(i))
   Next
   l = l + Val(txtQty(i))
   lblAllo = Format(l, "#####0")
   If l > Val(lblRqty) Then
      lblAllo.ForeColor = ES_RED
   Else
      lblAllo.ForeColor = vbBlack
   End If
   
End Sub

Public Sub GetCurrentAllocations()
   Dim RdoAll As rdoResultset
   Dim i As Integer
   Dim n As Integer
   Dim SPartRef As String
   Dim lSon(5, 3) As Long
   
   ClearBoxes
   SPartRef = Compress(cmbPrt)
   i = -1
   On Error Resume Next
   prg1.Visible = True
   prg1.Value = 5
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM RnalTable WHERE RAREF='" & SPartRef & "' " _
          & "AND RARUN=" & Val(cmbRun) & " "
   bSqlRows = GetDataSet(RdoAll)
   If bSqlRows Then
      With RdoAll
         Do Until .EOF
            i = i + 1
            If i > 4 Then Exit Do
            lSon(i, 0) = Format(0 + !RASO, "####0")
            lSon(i, 1) = Format(0 + !RAQTY, "####0")
            lblItm(i) = Format(!RASOITEM, "##0")
            lblRev(i) = "" & Trim(!RASOREV)
            .MoveNext
         Loop
         .Cancel
         prg1.Value = 50
      End With
      For n = 0 To i
         sSql = "SELECT SONUMBER,SOTYPE FROM SohdTable " _
                & "WHERE SONUMBER=" & lSon(n, 0) & " "
         bSqlRows = GetDataSet(RdoAll)
         If bSqlRows Then
            With RdoAll
               cmbSon(n) = "" & !SOTYPE & Format(lSon(n, 0), "00000")
               txtQty(n) = Format(0 + lSon(n, 1), "#####0")
               cmbItm(n) = lblItm(n) & Trim(lblRev(n))
               iIndex = n
               bGoodItem = GetThisItem()
               txtQty(n).Enabled = True
               .Cancel
            End With
            prg1.Value = 80
         End If
      Next
      iIndex = 0
   End If
   GetThisRun
   prg1.Value = 80
   Set RdoAll = Nothing
   prg1.Visible = False
   prg1.Value = 0
   Exit Sub
   
   DiaErr1:
   sProcName = "getcurrenta"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub AllocateItems()
   Dim bByte As Byte
   Dim i As Integer
   Dim bResponse As Integer
   Dim sMsg As String
   Dim SPartRef As String
   Dim vItems(5, 4) As Variant
   
   SPartRef = Compress(cmbPrt)
   For i = 0 To 4
      If Trim(cmbSon(i)) <> "" And Val(lblItm(i)) > 0 _
              And Val(txtQty(i)) > 0 Then
         bByte = True
         Exit For
      End If
   Next
   If Not bByte Then
      MsgBox "There Are Not Items To Allocate.", vbInformation, Caption
      Exit Sub
   End If
   If Not bDataChanged Then
      MsgBox "The Data Has Not Changed.", vbInformation, Caption
      Exit Sub
   End If
   sMsg = "Do You Want Replace Any Current " & vbCrLf _
          & "Allocations With These New Allocations?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      On Error GoTo RvsoAl1
      sSql = "DELETE FROM RnalTable WHERE RAREF='" & SPartRef & "' AND " _
             & "RARUN=" & Val(cmbRun) & " "
      RdoCon.Execute sSql, rdExecDirect
      For i = 0 To 4
         If Len(Trim(cmbSon(i))) > 0 And _
                Val(lblItm(i)) > 0 And Val(txtQty(i)) > 0 Then
            vItems(i, 0) = Right(cmbSon(i), 5)
            vItems(i, 1) = lblItm(i)
            vItems(i, 2) = "" & Trim(lblRev(i))
            vItems(i, 3) = txtQty(i)
         Else
            vItems(i, 3) = 0
         End If
      Next
      For i = 0 To 4
         If Val(vItems(i, 3)) > 0 Then
            sSql = "INSERT INTO RnalTable (RAREF,RARUN,RASO," _
                   & "RASOITEM,RASOREV,RAQTY) VALUES('" _
                   & SPartRef & "'," & Val(cmbRun) & "," _
                   & Val(vItems(i, 0)) & "," _
                   & Val(vItems(i, 1)) & ",'" _
                   & vItems(i, 2) & "'," _
                   & Val(vItems(i, 3)) & ")"
            RdoCon.Execute sSql, rdExecDirect
         End If
      Next
      MouseCursor 0
      MsgBox "Allocations Completed.", vbInformation, Caption
   Else
      CancelTrans
   End If
   Exit Sub
   
   RvsoAl1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume RvsoAl2
   RvsoAl2:
   MouseCursor 0
   MsgBox CurrError.Description & vbCrLf _
      & "Couldn't Complete The Allocations.", vbExclamation, Caption
   
End Sub

Public Sub Fillcombo()
   Dim RdoPcl As rdoResultset
   Dim sTempPart As String
   
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,PALEVEL,RUNREF," _
          & "RUNSTATUS FROM PartTable,RunsTable WHERE " _
          & "RUNREF=PARTREF AND (RUNSTATUS<>'CA' AND RUNSTATUS<>'CL')"
   bSqlRows = GetDataSet(RdoPcl)
   If bSqlRows Then
      With RdoPcl
         cmbPrt = "" & Trim(!PARTNUM)
         lblDsc = "" & Trim(!PADESC)
         Do Until .EOF
            If sTempPart <> Trim(!PARTNUM) Then
               'cmbPrt.AddItem "" & Trim(!PARTNUM)
               AddComboStr cmbPrt.hWnd, "" & Trim(!PARTNUM)
               sTempPart = Trim(!PARTNUM)
            End If
            .MoveNext
         Loop
      End With
      If cmbPrt.ListCount > 0 Then FillFormRuns
   Else
      MsgBox "No Matching Runs Recorded.", _
         vbExclamation, Caption
   End If
   On Error Resume Next
   Set RdoPcl = Nothing
   cmbPrt.SetFocus
   Exit Sub
   
   DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub ClearBoxes()
   Dim i As Integer
   lblRqty = ""
   For i = 0 To 4
      cmbSon(i) = ""
      cmbItm(i) = ""
      LblPrt(i) = ""
      lblPDsc(i) = ""
      lblItm(i) = ""
      lblRev(i) = ""
      lblQty(i) = ""
      txtQty(i) = ""
   Next
   
End Sub

Public Sub GetThisRun()
   Dim RdoStu As rdoResultset
   Dim SPartRef As String
   On Error GoTo DiaErr1
   SPartRef = Compress(cmbPrt)
   sSql = "SELECT RUNREF,RUNNO,RUNQTY FROM " _
          & "RunsTable WHERE RUNREF = '" & SPartRef & "' " _
          & "AND RUNNO=" & cmbRun & " "
   bSqlRows = GetDataSet(RdoStu, ES_FORWARD)
   If bSqlRows Then
      lblRqty = Format(RdoStu!RUNQTY, "######0")
   Else
      lblRqty = ""
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "getthisrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
