VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ActTree 
   BackColor       =   &H8000000C&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Structure"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "ActTree.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.TreeView tvw1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5741
      _Version        =   327682
      Indentation     =   353
      Style           =   7
      ImageList       =   "imlSmallIcons"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4320
      Top             =   3600
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3885
      FormDesignWidth =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chart Of Accounts Structure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
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
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ActTree.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ActTree.frx":0B8C
            Key             =   "cylinder"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ActTree.frx":10AE
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ActTree.frx":15D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ActTree.frx":18C6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ActTree"
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
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim tNode As Node



Private Sub Form_Activate()
   If bOnLoad = 1 Then
      FillActTree
      bOnLoad = 0
   End If
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   If Not bCancel Then Unload Me
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   If iBarOnTop Then
      Move MdiSect.Left + 4500, MdiSect.Top + 1900
   Else
      Move MdiSect.Left + 4500, MdiSect.Top + 1100
   End If
   bOnLoad = 1
   
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   WindowState = 1
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set ActTree = Nothing
   
End Sub


Private Sub tvw1_Collapse(ByVal Node As ComctlLib.Node)
   Node.Image = 1
   
End Sub

Private Sub tvw1_Expand(ByVal Node As ComctlLib.Node)
   Node.Image = 4
   
End Sub


Private Sub tvw1_NodeClick(ByVal Node As ComctlLib.Node)
   Dim sAccount As String
   sAccount = TrimAccount(Trim(Node.Text))
   If Len(sAccount) Then RecurTest Node.Index, sAccount
   
End Sub



Public Sub FillActTree()
   Dim A As Integer
   Dim i As Integer
   Dim rdoAct As ADODB.Recordset
   Dim sAccount As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   
   sSql = "SELECT COASSTACCT,COLIABACCT,COEQTYACCT," _
          & "COINCMACCT,COCOGSACCT,COEXPNACCT,COOINCACCT," _
          & "COOEXPACCT,COFDTXACCT FROM GlmsTable " _
          & "WHERE COACCTREC=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct)
   If bSqlRows Then
      With rdoAct
         For i = 0 To 7
            sAccount = "" & Trim(.Fields(i))
            If Len(sAccount) Then
               Set tNode = tvw1.Nodes.Add(, , , Trim(.Fields(i)), 1)
               A = tNode.Index
               '  sAccount = .Fields(i)
               '  RecurTest a, sAccount
            End If
         Next
         sAccount = "" & Trim(.Fields(i))
         If Len(sAccount) Then
            Set tNode = tvw1.Nodes.Add(, , , .Fields(i), 1)
            A = tNode.Index
            ' sAccount = Compress(.Fields(i))
            ' RecurTest a, sAccount
         End If
         .Cancel
      End With
   End If
   MouseCursor 0
   Set rdoAct = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillacttr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub










'This routine is a test of using recursion as a
'method of gather accounts.  Make the code
'smaller and eliminates the control of depth
'
'The only errors here should be duplicate keys
'in that case we bail.
'
'Changed to load on demand 6/30/99

Public Sub RecurTest(iNode As Integer, sMaster As String)
   Dim i As Integer
   Dim RdoAct1 As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillAccounts '" & sMaster & "',0"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct1)
   If bSqlRows Then
      On Error Resume Next
      With RdoAct1
         Do Until .EOF
            Set tNode = tvw1.Nodes.Add(iNode, tvwChild, Trim(!GLACCTNO) & str(i), Trim(!GLACCTNO) & ": " & Left$(!GLDESCR, 30))
            If Err > 0 Then Exit Do
            i = tNode.Index
            RecurTest i, Trim$(!GLACCTREF)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoAct1 = Nothing
   Exit Sub
   
DiaErr1:
   Resume DiaErr2
DiaErr2:
   On Error GoTo 0
   
End Sub

Public Function TrimAccount(sAccount As String) As String
   Dim A As Integer
   A = Len(sAccount)
   If A > 1 Then
      A = InStr(sAccount, Chr$(58))
      If A > 0 Then
         sAccount = Left$(sAccount, A - 1)
      End If
      sAccount = Compress(sAccount)
   End If
   TrimAccount = sAccount
   
End Function
