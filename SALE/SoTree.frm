VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form SoTree
   BorderStyle = 3 'Fixed Dialog
   Caption = "Sales Orders"
   ClientHeight = 3972
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 3516
   Icon = "SoTree.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 3972
   ScaleWidth = 3516
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
      Style = 7
      ImageList = "imlSmallIcons"
      Appearance = 1
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 3000
      Top = 3600
      _Version = 196615
      _ExtentX = 593
      _ExtentY = 593
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3972
      FormDesignWidth = 3516
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
      NumListImages = 5
      BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "SoTree.frx":030A
      Key = ""
      EndProperty
      BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "SoTree.frx":05CC
      Key = "cylinder"
      EndProperty
      BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "SoTree.frx":0AEE
      Key = "leaf"
      EndProperty
      BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "SoTree.frx":1010
      Key = ""
      EndProperty
      BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "SoTree.frx":1306
      Key = ""
      EndProperty
      EndProperty
   End
End
Attribute VB_Name = "SoTree"
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

Private Sub Form_Load()
   SetFormSize Me
   If iBarOnTop Then
      Move MdiSect.Left + 3500, MdiSect.Top + 1900
   Else
      Move MdiSect.Left + 5500, MdiSect.Top + 1100
   End If
   FillTree
   
End Sub



Private Sub FillTree()
   Dim RdoCst As rdoResultset
   Dim iList As Integer
   Dim sCust As String
   Dim sSONum As String
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "SELECT DISTINCT SOCUST,SOTYPE,SONUMBER,SODATE,CUREF,CUNICKNAME FROM " _
          & "SohdTable,CustTable WHERE SOCUST =CUREF ORDER BY SOCUST"
   bSqlRows = GetDataSet(RdoCst, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoCst
         Do Until .EOF
            If sCust <> Trim(!CUREF) Then
               Set tNode = tvw1.Nodes.Add(, , , Trim(!CUNICKNAME), 1)
               sCust = Trim(!CUREF)
               iList = tNode.Index
            End If
            sSONum = Trim(!SOTYPE) & Format(!SONUMBER, "000000") _
                     & " " & Format(!SODATE, "mm/dd/yy")
            Set tNode = tvw1.Nodes.Add(iList, tvwChild, , sSONum)
            .MoveNext
         Loop
         ClearResultSet RdoCst
      End With
   End If
   MouseCursor 0
   Set RdoCst = Nothing
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
   Set SoTree = Nothing
   
End Sub


Private Sub tvw1_Collapse(ByVal Node As ComctlLib.Node)
   Node.Image = 1
   
End Sub

Private Sub tvw1_Expand(ByVal Node As ComctlLib.Node)
   Node.Image = 4
   
End Sub


Private Sub tvw1_NodeClick(ByVal Node As ComctlLib.Node)
   If Val(Mid(Node.Text, 2, 6)) > 0 Then
      bCancel = True
      MdiSect.ActiveForm.cmbSon = Mid(Node.Text, 2, 6)
      MdiSect.ActiveForm.optSvw.Value = vbChecked
      Unload Me
   End If
   
End Sub
