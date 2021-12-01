VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form RteTree 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routing"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "RteTree.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1155
   End
   Begin ComctlLib.TreeView tvw1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4048
      _Version        =   327682
      Style           =   7
      ImageList       =   "imlSmallIcons"
      Appearance      =   1
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   2880
      Top             =   3600
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3975
      FormDesignWidth =   3495
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Extended Part Description"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   3255
   End
   Begin ComctlLib.ImageList imlSmallIcons 
      Left            =   0
      Top             =   3480
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
            Picture         =   "RteTree.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RteTree.frx":05CC
            Key             =   "cylinder"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RteTree.frx":0AEE
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RteTree.frx":1010
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RteTree.frx":1306
            Key             =   "smlBook"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RteTree.frx":1968
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "RteTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bCancel As Byte
Dim tNode As Node

Dim sParts(500) As String

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub Form_Deactivate()
   If Not bCancel Then Unload Me
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   If iBarOnTop Then
      Move MDISect.Left + 2200, MDISect.Top + 1900
   Else
      Move MDISect.Left + 5000, MDISect.Top + 1100
   End If
   FillTree
   
End Sub



Private Sub FillTree()
   Dim RdoPrt As ADODB.Recordset
   Dim A As Integer
   Dim iList As Integer
   Dim sRout1 As String
   Dim sRout2 As String
   Dim sPart As String
   sRout1 = MDISect.ActiveForm.cmbRte _
            & " " & Trim(MDISect.ActiveForm.txtDsc)
   sRout2 = Compress(MDISect.ActiveForm.cmbRte)
   
   On Error GoTo DiaErr1
   Set tNode = tvw1.Nodes.Add(, , , sRout1, 1)
   iList = tNode.Index
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL FROM " _
          & "PartTable WHERE PAROUTING='" & sRout2 & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
   If bSqlRows Then
      On Error Resume Next
      With RdoPrt
         Do Until .EOF
            A = A + 1
            sParts(A) = Trim(!PartRef)
            sPart = "" & Trim(!PartNum) & " " _
                    & Trim(!PADESC) & " Level:" & str(!PALEVEL)
            Set tNode = tvw1.Nodes.Add(iList, tvwChild, , sPart)
            .MoveNext
         Loop
         ClearResultSet RdoPrt
      End With
   End If
   Set RdoPrt = Nothing
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
   Set RteTree = Nothing
   
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



Private Sub GetPartExt(iIndex As Integer)
   Dim RdoPrt As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PAEXTDESC FROM " _
          & "PartTable WHERE PARTREF='" & sParts(iIndex) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_STATIC)
   If bSqlRows Then
      lblDsc = "" & Trim(RdoPrt!PAEXTDESC)
      ClearResultSet RdoPrt
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
DiaErr1:
   On Error GoTo 0
   
End Sub
