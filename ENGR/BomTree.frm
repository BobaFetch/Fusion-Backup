VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ViewBomTree 
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.TreeView tvw1 
      Height          =   3012
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3852
      _ExtentX        =   6800
      _ExtentY        =   5318
      _Version        =   327682
      Indentation     =   529
      Style           =   7
      ImageList       =   "imlSmallIcons"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdEdit 
      Cancel          =   -1  'True
      Caption         =   "&Edit"
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   1440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1155
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   4200
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
   Begin ComctlLib.ImageList imlSmallIcons 
      Left            =   600
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BomTree.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BomTree.frx":05CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   3855
   End
End
Attribute VB_Name = "ViewBomTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Major overhaul 7/16/03
'10/24/03 Added "key:" to TreeView Key work around a bug
'8/20/06 Corrected Clear and opening bitmap
Option Explicit
'Dim RdoQry As rdoQuery
Dim AdoCmdObj As ADODB.Command

Dim bCancel As Byte
Dim iCurrIdx As Integer
Dim iKey As Integer
Dim tNode As Node

Dim sCurrPart As String

Dim sBillParts(700, 4) As String

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
      Move MDISect.Left + 2200, MDISect.Top + 1900
   Else
      Move MDISect.Left + 5000, MDISect.Top + 1100
   End If
   sSql = "SELECT BMASSYPART,BMPARTREF,BMPARTNUM,BMREV," _
          & "BMSEQUENCE,BMQTYREQD,BMUNITS," _
          & "PARTREF,PARTNUM,PADESC,PALEVEL FROM BmplTable,PartTable " _
          & "WHERE BMPARTREF=PARTREF AND (BMASSYPART= ? AND " _
          & "BMREV= ? ) ORDER BY BMSEQUENCE,BMPARTREF "
   
   Set AdoCmdObj = New ADODB.Command
   AdoCmdObj.CommandText = sSql
   
   Dim prmAssPrt As ADODB.Parameter
   Set prmAssPrt = New ADODB.Parameter
   prmAssPrt.Type = adChar
   prmAssPrt.Size = 30
   AdoCmdObj.Parameters.Append prmAssPrt
   
   Dim prmBMRev As ADODB.Parameter
   Set prmBMRev = New ADODB.Parameter
   prmBMRev.Type = adChar
   prmBMRev.Size = 4
   AdoCmdObj.Parameters.Append prmBMRev
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   
   FillTree
   
End Sub



Private Sub FillTree()
   Dim RdoBom As ADODB.Recordset
   Dim A As Integer
   Dim iList As Integer
   Dim sPl1 As String
   Dim sPl2 As String
   Dim sRev As String
   
   iKey = 0
   tvw1.Nodes.Clear
   Erase sBillParts
   sRev = MDISect.ActiveForm.cmbRev
   sPl1 = MDISect.ActiveForm.cmbPls _
          & " " & Trim(MDISect.ActiveForm.lblDsc) & " Rev: " & sRev
   sPl2 = Compress(MDISect.ActiveForm.cmbPls)
   On Error Resume Next
   Set tNode = tvw1.Nodes.Add(, , "0" & sPl2, sPl1, 1)
   sBillParts(tNode.Index, 0) = MDISect.ActiveForm.cmbPls
   iList = tNode.Index
   sSql = "SELECT BMASSYPART,BMPARTREF,BMPARTNUM,BMREV," _
          & "BMSEQUENCE,BMQTYREQD,BMUNITS," _
          & "PARTREF,PARTNUM,PADESC,PALEVEL FROM BmplTable,PartTable " _
          & "WHERE BMPARTREF=PARTREF AND (BMASSYPART='" & sPl2 & "' AND " _
          & "BMREV='" & sRev & "') ORDER BY BMSEQUENCE,BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_STATIC)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iKey = iKey + 1
            Set tNode = tvw1.Nodes.Add(, , str$(iKey) & Trim(!BMASSYPART) & Trim(!BMPARTREF), "" & Trim(!PartNum) & " Rev: " & Trim(!BMREV))
            sBillParts(tNode.Index, 0) = "" & Trim(!PartRef)
            sBillParts(tNode.Index, 1) = "" & Trim(!BMREV)
            sBillParts(tNode.Index, 2) = "" & Trim(!PADESC) _
                       & " Qty " & Format$(!BMQTYREQD, ES_QuantityDataFormat) & Trim(!BMUNITS)
            sBillParts(tNode.Index, 3) = str$(!PALEVEL)
            NextBillLevel tNode.Index, sBillParts(tNode.Index, 0), sBillParts(tNode.Index, 1)
            .MoveNext
         Loop
         ClearResultSet RdoBom
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
   Set AdoCmdObj = Nothing
   Set ViewBomTree = Nothing
   
End Sub


Private Sub tvw1_Click()
   lblDsc = sBillParts(iCurrIdx, 2)
   
End Sub

Private Sub tvw1_Collapse(ByVal Node As ComctlLib.Node)
   Node.Image = 1
   iCurrIdx = Node.Index
   
End Sub

Private Sub tvw1_Expand(ByVal Node As ComctlLib.Node)
   Node.Image = 2
   iCurrIdx = Node.Index
   
End Sub

Private Sub tvw1_NodeClick(ByVal Node As ComctlLib.Node)
   sCurrPart = Compress(Node.Text)
   iCurrIdx = Node.Index
   'If Len(sCurrPart) Then NextBillLevel Node.Index, sBillParts(Node.Index, 0), sBillParts(Node.Index, 1)
   
End Sub



Private Sub NextBillLevel(iNode As Integer, sPartNumber As String, sRev As String)
   Dim RdoBm1 As ADODB.Recordset
   Dim iList As Integer
   
   'On Error GoTo DiaErr1
   On Error Resume Next
   
   AdoCmdObj.Parameters(0).value = sPartNumber
   AdoCmdObj.Parameters(1).value = sRev
'   RdoQry(0) = sPartNumber
'   RdoQry(1) = sRev
   bSqlRows = clsADOCon.GetQuerySet(RdoBm1, AdoCmdObj, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBm1
         Do Until .EOF
            iKey = iKey + 1
            Set tNode = tvw1.Nodes.Add(iNode, tvwChild, str$(iKey) & Trim(!BMASSYPART) & Trim(!PartRef), Trim(!PartNum) & " Rev: " & !BMREV)
            If Err > 0 Then Exit Do
            iList = tNode.Index
            sBillParts(tNode.Index, 0) = "" & Trim(!PartRef)
            sBillParts(tNode.Index, 1) = "" & Trim(!BMREV)
            sBillParts(tNode.Index, 2) = "" & Trim(!PADESC) _
                       & " Qty " & Format$(!BMQTYREQD, ES_QuantityDataFormat) & Trim(!BMUNITS)
            sBillParts(tNode.Index, 3) = str$(!PALEVEL)
            NextBillLevel iList, Trim(!PartRef), "" & Trim(!BMREV)
            .MoveNext
         Loop
         ClearResultSet RdoBm1
      End With
   End If
   Set RdoBm1 = Nothing
   Exit Sub
   
DiaErr1:
   Resume DiaErr2
DiaErr2:
   On Error GoTo 0
   
End Sub
