VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ViewBom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parts List"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   Icon            =   "BomTree.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEdit 
      Cancel          =   -1  'True
      Caption         =   "&Edit"
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   1440
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1155
   End
   Begin ComctlLib.TreeView tvw1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Click Items For Detail"
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5318
      _Version        =   327682
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imlSmallIcons"
      Appearance      =   1
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   4080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4290
      FormDesignWidth =   4110
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   3855
   End
   Begin ComctlLib.ImageList imlSmallIcons 
      Left            =   0
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BomTree.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BomTree.frx":05CC
            Key             =   "cylinder"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BomTree.frx":0AEE
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BomTree.frx":1010
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BomTree.frx":1306
            Key             =   "smlBook"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BomTree.frx":1968
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ViewBom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Major overhaul 7/16/03
'10/24/03 Added "key:" to TreeView Key work around to bug
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim AdoParameter2 As ADODB.Parameter

Dim bCancel As Byte
Dim iCurrIdx As Integer
Dim tNode As Node

Dim sCurrPart As String

Dim sBillParts(700, 3) As String

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdEdit_Click()
   If iCurrIdx > 0 Then
      MsgBox sBillParts(iCurrIdx, 0)
   End If
   
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
      Move MdiSect.Left + 2200, MdiSect.Top + 1900
   Else
      Move MdiSect.Left + 5000, MdiSect.Top + 1100
   End If
   sSql = "SELECT BMASSYPART,BMPARTREF,BMPARTNUM,BMREV," _
          & "BMSEQUENCE,BMQTYREQD,BMUNITS," _
          & "PARTREF,PARTNUM,PADESC FROM BmplTable,PartTable " _
          & "WHERE BMPARTREF=PARTREF AND (BMASSYPART= ? AND " _
          & "BMREV= ? ) ORDER BY BMSEQUENCE,BMPARTREF "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   AdoQry.parameters.Append AdoParameter1
   Set AdoParameter2 = New ADODB.Parameter
   AdoParameter2.Type = adChar
   AdoParameter2.SIZE = 2
   AdoQry.parameters.Append AdoParameter2
   
   FillTree
   
End Sub



Private Sub FillTree()
   Dim RdoBom As ADODB.Recordset
   Dim A As Integer
   Dim iList As Integer
   Dim sPl1 As String
   Dim sPl2 As String
   Dim sRev As String
   
   sRev = MdiSect.ActiveForm.cmbRev
   sPl1 = MdiSect.ActiveForm.cmbPls _
          & " " & Trim(MdiSect.ActiveForm.lblDsc) & " Rev: " & sRev
   sPl2 = Compress(MdiSect.ActiveForm.cmbPls)
   On Error GoTo DiaErr1
   Set tNode = tvw1.Nodes.Add(, , "key:" & sPl2, sPl1, 1)
   iList = tNode.Index
   sSql = "SELECT BMASSYPART,BMPARTREF,BMPARTNUM,BMREV," _
          & "BMSEQUENCE,BMQTYREQD,BMUNITS," _
          & "PARTREF,PARTNUM,PADESC FROM BmplTable,PartTable " _
          & "WHERE BMPARTREF=PARTREF AND (BMASSYPART='" & sPl2 & "' AND " _
          & "BMREV='" & sRev & "') ORDER BY BMSEQUENCE,BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_STATIC)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            Set tNode = tvw1.Nodes.Add(, , "key:" & Trim(!BMASSYPART) & Trim(!PartRef), "" & Trim(!PARTNUM) & " Rev: " & Trim(!BMREV))
            sBillParts(tNode.Index, 0) = "" & Trim(!PartRef)
            sBillParts(tNode.Index, 1) = "" & Trim(!BMREV)
            sBillParts(tNode.Index, 2) = "" & Trim(!PADESC) _
                       & " Qty " & Format$(!BMQTYREQD, ES_QuantityDataFormat) & Trim(!BMUNITS)
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
   Set AdoParameter1 = Nothing
   Set AdoParameter2 = Nothing
   Set AdoQry = Nothing
   Set ViewBom = Nothing
   
End Sub


Private Sub tvw1_Click()
   lblDsc = sBillParts(iCurrIdx, 2)
   
End Sub

Private Sub tvw1_Collapse(ByVal Node As ComctlLib.Node)
   Node.Image = 1
   iCurrIdx = Node.Index
   
End Sub

Private Sub tvw1_Expand(ByVal Node As ComctlLib.Node)
   Node.Image = 4
   iCurrIdx = Node.Index
   
End Sub

Private Sub tvw1_NodeClick(ByVal Node As ComctlLib.Node)
   sCurrPart = Compress(Node.Text)
   iCurrIdx = Node.Index
   If Len(sCurrPart) Then NextBillLevel Node.Index, sBillParts(Node.Index, 0), sBillParts(Node.Index, 1)
   
End Sub



Private Sub NextBillLevel(iNode As Integer, sPartNumber As String, sRev As String)
   Dim RdoBm1 As ADODB.Recordset
   Dim iList As Integer
   
   On Error GoTo DiaErr1
   AdoQry.parameters(0).Value = sPartNumber
   AdoQry.parameters(1).Value = sRev
   bSqlRows = clsADOCon.GetQuerySet(RdoBm1, AdoQry, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBm1
         Do Until .EOF
            Set tNode = tvw1.Nodes.Add(iNode, tvwChild, "key:" & Trim(!BMASSYPART) & Trim(!PartRef), Trim(!PARTNUM) & " Rev: " & !BMREV)
            If Err > 0 Then Exit Do
            iList = tNode.Index
            sBillParts(tNode.Index, 0) = "" & Trim(!PartRef)
            sBillParts(tNode.Index, 1) = "" & Trim(!BMREV)
            sBillParts(tNode.Index, 2) = "" & Trim(!PADESC) _
                       & " Qty " & Format$(!BMQTYREQD, ES_QuantityDataFormat) & Trim(!BMUNITS)
            NextBillLevel iList, Trim(!PartRef), "" & Trim(!BMREV)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoBm1 = Nothing
   Exit Sub
   
DiaErr1:
   Resume DiaErr2
DiaErr2:
   On Error GoTo 0
   
End Sub
