VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form EstTree
   BorderStyle = 3 'Fixed Dialog
   Caption = "Estimating Bill Of Material"
   ClientHeight = 4320
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 4368
   Icon = "EstTree.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 4320
   ScaleWidth = 4368
   ShowInTaskbar = 0 'False
   Begin VB.CheckBox optCost
      Caption = "Check1"
      Height = 195
      Left = 720
      TabIndex = 3
      Top = 4080
      Visible = 0 'False
      Width = 495
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 1680
      TabIndex = 1
      TabStop = 0 'False
      Top = 3840
      Width = 1155
   End
   Begin ComctlLib.TreeView tvw1
      Height = 2295
      Left = 120
      TabIndex = 0
      ToolTipText = "Right Click Mouse To Show Costs"
      Top = 480
      Width = 4095
      _ExtentX = 7218
      _ExtentY = 4043
      _Version = 327682
      LabelEdit = 1
      LineStyle = 1
      Style = 7
      ImageList = "imlSmallIcons"
      Appearance = 1
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 2880
      Top = 3840
      _Version = 196615
      _ExtentX = 593
      _ExtentY = 593
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 4320
      FormDesignWidth = 4368
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Click The Item And Then Right Click For Costs"
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 4
      Top = 120
      Width = 4095
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Caption = "Extended Part Description"
      Height = 855
      Left = 120
      TabIndex = 2
      Top = 2880
      Width = 4095
   End
   Begin ComctlLib.ImageList imlSmallIcons
      Left = 120
      Top = 3840
      _ExtentX = 995
      _ExtentY = 995
      BackColor = -2147483643
      ImageWidth = 13
      ImageHeight = 13
      MaskColor = 12632256
      _Version = 327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7}
      NumListImages = 6
      BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "EstTree.frx":030A
      Key = ""
      EndProperty
      BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "EstTree.frx":05CC
      Key = "cylinder"
      EndProperty
      BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "EstTree.frx":0AEE
      Key = "leaf"
      EndProperty
      BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "EstTree.frx":1010
      Key = ""
      EndProperty
      BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "EstTree.frx":1306
      Key = "smlBook"
      EndProperty
      BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "EstTree.frx":1968
      Key = ""
      EndProperty
      EndProperty
   End
End
Attribute VB_Name = "EstTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2006) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim tNode As Node
Dim tOpen As Node

Dim sEstPart As String



Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub Form_Activate()
   If bOnLoad = 1 Then Show
   bOnLoad = 0
   If optCost.Value = vbChecked Then
      bCancel = 0
      Unload diaEsBcs
      ' optCost.Value = vbUnchecked
   End If
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   If bCancel = 0 Then Unload Me
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   
   If iBarOnTop Then
      Move MdiSect.Left + 2200, MdiSect.Top + 1900
   Else
      Move MdiSect.Left + 5000, MdiSect.Top + 1100
   End If
   bOnLoad = 1
   FillTree
   
End Sub



Private Sub FillTree()
   Dim RdoEst As rdoResultset
   Dim a As Integer
   Dim iList As Integer
   Dim sPart As String
   sEstPart = sPart
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE " _
          & "PARTREF='" & sPart & "'"
   bSqlRows = GetDataSet(RdoEst, ES_FORWARD)
   If bSqlRows Then
      With RdoEst
         Set tNode = tvw1.Nodes.Add(, , , "" & Trim(!PARTNUM), 1)
         iList = tNode.Index
         GetNextBillLevel iList, sPart
         ClearResultSet RdoEst
      End With
   End If
   
   Set RdoEst = Nothing
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   sProcName = "filltree"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If optCost.Value = vbChecked Then Unload diaEsBcs
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set EstTree = Nothing
   
End Sub


Private Sub optCost_Click()
   'Cost is showing
   
End Sub

Private Sub tvw1_Collapse(ByVal Node As ComctlLib.Node)
   Node.Image = 1
   lblDsc = "Part Description"
   
End Sub

Private Sub tvw1_DblClick()
   ' diaEsBcs.Show
   
End Sub


Private Sub tvw1_Expand(ByVal Node As ComctlLib.Node)
   Node.Image = 4
   lblDsc = "Part Description"
   
End Sub

Private Sub tvw1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      On Error GoTo ProcErr1
      If sEstPart <> Compress(tOpen) Then
         If Not IsNull(tOpen) Then
            bCancel = 1
            optCost.Value = vbChecked
            diaEsBcs.lblPart = tOpen
            diaEsBcs.Show
         End If
      Else
         MsgBox "Cannot Edit Costs Of Estimated Part.", vbInformation, _
            Caption
      End If
   End If
   Exit Sub
   ProcErr1:
   On Error GoTo 0
   
End Sub

Private Sub tvw1_NodeClick(ByVal Node As ComctlLib.Node)
   Set tOpen = Node
   If Node.Index > 0 Then
      GetPartExt Node.Text
   Else
      lblDsc = ""
   End If
   
End Sub



Private Sub GetPartExt(sParts As String)
   Dim RdoPrt As rdoResultset
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PADESC,PAEXTDESC FROM " _
          & "PartTable WHERE PARTREF='" & Compress(sParts) & "'"
   bSqlRows = GetDataSet(RdoPrt, ES_STATIC)
   If bSqlRows Then
      lblDsc = "Desc:       " & Trim(RdoPrt!PADESC) & vbCr _
               & "Extended: " & Trim(RdoPrt!PAEXTDESC)
      ClearResultSet RdoPrt
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
   DiaErr1:
   On Error GoTo 0
   
End Sub

Private Sub GetNextBillLevel(iNode As Integer, sBomPart As String)
   Dim RdoBom As rdoResultset
   Dim iList As Integer
   Dim sBillPart As String
   
   On Error GoTo DiaErr1
   sSql = "SELECT BMASSYPART,BMPARTREF,PARTREF,PARTNUM FROM " _
          & "BmplTable,PartTable WHERE (PARTREF=BMASSYPART " _
          & "AND BMASSYPART='" & sBomPart & "')"
   bSqlRows = GetDataSet(RdoBom, ES_FORWARD)
   If bSqlRows Then
      With RdoBom
         Do Until .EOF
            sProcName = "getnextbill"
            sBillPart = GetBillPart("" & Trim(!BMPARTREF))
            If sBillPart <> "" Then
               Set tNode = tvw1.Nodes.Add(iNode, tvwChild, Trim(!BMPARTREF), sBillPart)
               iList = tNode.Index
               GetNextBillLevel iList, Trim(!BMPARTREF)
            End If
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "getnextbill"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetBillPart(sPartNo As String)
   Dim RdoBprt As rdoResultset
   
   On Error GoTo DiaErr1
   sProcName = "getbillpart"
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable " _
          & "WHERE PARTREF='" & sPartNo & "'"
   bSqlRows = GetDataSet(RdoBprt, ES_FORWARD)
   If bSqlRows Then
      With RdoBprt
         GetBillPart = "" & Trim(!PARTNUM)
         ClearResultSet RdoBprt
      End With
   Else
      GetBillPart = ""
   End If
   Set RdoBprt = Nothing
   Exit Function
   
   DiaErr1:
   sProcName = "getbillpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
