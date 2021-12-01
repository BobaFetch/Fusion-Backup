VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form VewBomTree
   BackColor = &H8000000C&
   BorderStyle = 3 'Fixed Dialog
   Caption = "Parts List"
   ClientHeight = 4035
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 3495
   Icon = "BomTree.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 4035
   ScaleWidth = 3495
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
      Height = 2295
      Left = 120
      TabIndex = 0
      Top = 120
      Width = 3255
      _ExtentX = 5741
      _ExtentY = 4048
      _Version = 327682
      Style = 7
      ImageList = "imlSmallIcons"
      Appearance = 1
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 2880
      Top = 3600
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 4035
      FormDesignWidth = 3495
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Caption = "Extended Part Description"
      Height = 855
      Left = 120
      TabIndex = 2
      Top = 2520
      Width = 3255
   End
   Begin ComctlLib.ImageList imlSmallIcons
      Left = 0
      Top = 3480
      _ExtentX = 1005
      _ExtentY = 1005
      BackColor = -2147483643
      ImageWidth = 13
      ImageHeight = 13
      MaskColor = 12632256
      _Version = 327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7}
      NumListImages = 6
      BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "BomTree.frx":030A
      Key = ""
      EndProperty
      BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "BomTree.frx":05CC
      Key = "cylinder"
      EndProperty
      BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "BomTree.frx":0AEE
      Key = "leaf"
      EndProperty
      BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "BomTree.frx":1010
      Key = ""
      EndProperty
      BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "BomTree.frx":1306
      Key = "smlBook"
      EndProperty
      BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
      Picture = "BomTree.frx":1968
      Key = ""
      EndProperty
      EndProperty
   End
End
Attribute VB_Name = "VewBomTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' vewBOMTree - View BOM structure
'
'
' Created: (cjs)
' Revions:
'   07/11/02 (nth) Ported from EsiEngr to be revised and used in EsiFina
'
'********************************************************************************

Option Explicit
Dim bCancel As Byte
Dim tNode As Node

Dim sParts(500) As String

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
      Move MdiSect.Left + 2200, MdiSect.Top + 1900
   Else
      Move MdiSect.Left + 5000, MdiSect.Top + 1100
   End If
   FillTree
End Sub

Public Sub FillTree()
   Dim RdoBom As rdoResultset
   Dim a As Integer
   Dim i As Integer
   Dim sPl1 As String
   Dim sPl2 As String
   Dim sRev As String
   Dim sPart As String
   
   '************************
   ' Changed from engineering code to connect to finanace
   ' 07/11/02 (nth)
   '************************
   sRev = MdiSect.ActiveForm.lblBOM
   sPl1 = MdiSect.ActiveForm.cmbprt _
          & " " & Trim(MdiSect.ActiveForm.lblDsc) & " Rev: " & sRev
   sPl2 = Compress(MdiSect.ActiveForm.cmbprt)
   
   
   
   On Error GoTo DiaErr1
   Set tNode = tvw1.Nodes.Add(, , , sPl1, 1)
   i = tNode.Index
   sSql = "SELECT BMASSYPART,BMPARTREF,BMPARTNUM,BMREV," _
          & "BMSEQUENCE,BMQTYREQD,BMUNITS," _
          & "PARTREF,PARTNUM,PADESC FROM BmplTable,PartTable " _
          & "WHERE BMPARTREF=PARTREF AND (BMASSYPART='" & sPl2 & "' AND " _
          & "BMREV='" & sRev & "') ORDER BY BMSEQUENCE,BMPARTREF "
   bSqlRows = GetDataSet(RdoBom, ES_STATIC)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            a = a + 1
            sParts(a) = Trim(!PARTREF)
            sPart = "" & Trim(!PARTNUM) & " " _
                    & Trim(!PADESC) & " Qty: " & Format(!BMQTYREQD, "####0.000") _
                    & " " & Trim(!BMUNITS)
            Set tNode = tvw1.Nodes.Add(i, tvwChild, , sPart)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoBom = Nothing
   Exit Sub
   
   DiaErr1:
   Resume DiaErr2
   DiaErr2:
   On Error GoTo 0
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set VewBomTree = Nothing
End Sub


Private Sub tvw1_Collapse(ByVal Node As ComctlLib.Node)
   Node.Image = 1
   lblDsc = "Extended Part Description"
   
End Sub

Private Sub tvw1_Expand(ByVal Node As ComctlLib.Node)
   Node.Image = 4
   lblDsc = "Extended Part Description"
   
End Sub

Private Sub tvw1_NodeClick(ByVal Node As ComctlLib.Node)
   If Node.Index > 1 Then
      GetPartExt Node.Index - 1
   Else
      lblDsc = ""
   End If
   
End Sub

Public Sub GetPartExt(iIndex As Integer)
   Dim rdoPrt As rdoResultset
   On Error Resume Next
   sSql = "SELECT PARTREF,PAEXTDESC FROM " _
          & "PartTable WHERE PARTREF='" & sParts(iIndex) & "'"
   bSqlRows = GetDataSet(rdoPrt, ES_STATIC)
   If bSqlRows Then
      lblDsc = "" & Trim(rdoPrt!PAEXTDESC)
      rdoPrt.Cancel
   End If
   Set rdoPrt = Nothing
   
End Sub
