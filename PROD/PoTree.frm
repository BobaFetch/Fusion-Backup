VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form PurcPOtree
   BorderStyle = 3 'Fixed Dialog
   Caption = "Purchase Orders"
   ClientHeight = 3972
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 3492
   Icon = "PoTree.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 3972
   ScaleWidth = 3492
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 315
      Left = 1200
      TabIndex = 1
      TabStop = 0 'False
      Top = 3600
      Width = 1155
   End
   Begin ComctlLib.TreeView tvw1
      Height = 3375
      Left = 120
      TabIndex = 0
      Top = 120
      Width = 3255
      _ExtentX = 5736
      _ExtentY = 5948
      _Version = 327682
      Indentation = 706
      Style = 7
      ImageList = "imlSmallIcons"
      Appearance = 1
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 2880
      Top = 3600
      _Version = 196615
      _ExtentX = 593
      _ExtentY = 593
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3972
      FormDesignWidth = 3492
   End
   Begin ComctlLib.ImageList imlSmallIcons
      Left = 0
      Top = 3480
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
      Picture = "PoTree.frx":030A
      Key = ""
      EndProperty
      BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "PoTree.frx":05CC
      Key = "cylinder"
      EndProperty
      BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "PoTree.frx":0AEE
      Key = "leaf"
      EndProperty
      BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "PoTree.frx":1010
      Key = ""
      EndProperty
      BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "PoTree.frx":1306
      Key = "smlBook"
      EndProperty
      BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "PoTree.frx":1968
      Key = ""
      EndProperty
      EndProperty
   End
End
Attribute VB_Name = "PurcPOtree"
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
Dim tNode As Node

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   If Not bCancel Then Unload Me
   
End Sub

Private Sub Form_Initialize()
   BackColor = ES_ViewBackColor
   
End Sub


Private Sub Form_Load()
   SetFormSize Me
   If iBarOnTop Then
      Move MDISect.Left + 2200, MDISect.Top + 1900
   Else
      Move MDISect.Left + 5000, MDISect.Top + 1100
   End If
   FillTree
   Show
   
End Sub



Private Sub FillTree()
   Dim RdoVed As rdoResultset
   Dim iList As Integer
   Dim sVendor As String
   Dim sPONum As String
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT POVENDOR,PONUMBER,PODATE,VEREF,VENICKNAME FROM " _
          & "PohdTable,VndrTable WHERE POVENDOR =VEREF"
   bSqlRows = GetDataSet(RdoVed)
   If bSqlRows Then
      On Error Resume Next
      With RdoVed
         Do Until RdoVed.EOF
            If sVendor <> Trim(!VEREF) Then
               Set tNode = tvw1.Nodes.Add(, , , Trim(!VENICKNAME), 1)
               sVendor = Trim(!VEREF)
               iList = tNode.Index
            End If
            sPONum = Format(!PONUMBER, "000000") & " " & Format(!PODATE, "mm/dd/yy")
            Set tNode = tvw1.Nodes.Add(iList, tvwChild, , sPONum)
            .MoveNext
         Loop
         ClearResultSet RdoVed
      End With
   End If
   Set RdoVed = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "filltree"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set PurcPOtree = Nothing
   
End Sub


Private Sub tvw1_Collapse(ByVal Node As ComctlLib.Node)
   Node.Image = 1
   
End Sub

Private Sub tvw1_Expand(ByVal Node As ComctlLib.Node)
   Node.Image = 4
   
End Sub

Private Sub tvw1_NodeClick(ByVal Node As ComctlLib.Node)
   If Val(Left(Node.Text, 6)) > 0 Then
      bCancel = True
      MDISect.ActiveForm.cmbPon = Left(Node.Text, 6)
      MDISect.ActiveForm.optPvw.Value = vbChecked
      Unload Me
   End If
   
End Sub
