VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1
   BorderStyle = 3 'Fixed Dialog
   Caption = "Form1"
   ClientHeight = 4344
   ClientLeft = 36
   ClientTop = 324
   ClientWidth = 5388
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 4344
   ScaleWidth = 5388
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton Command1
      Caption = "Command1"
      Height = 492
      Left = 3360
      TabIndex = 1
      Top = 3480
      Width = 1092
   End
   Begin MSComctlLib.TreeView TreeView1
      Height = 3045
      Left = 960
      TabIndex = 0
      Top = 120
      Width = 3855
      _ExtentX = 6795
      _ExtentY = 5376
      _Version = 393217
      Indentation = 529
      Style = 7
      Appearance = 1
   End
   Begin MSComctlLib.ImageList ImageList1
      Left = 600
      Top = 3480
      _ExtentX = 804
      _ExtentY = 804
      BackColor = -2147483643
      ImageWidth = 16
      ImageHeight = 16
      MaskColor = 12632256
      _Version = 393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628}
      NumListImages = 2
      BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628}
      Picture = "Form1.frx":0000
      Key = ""
      EndProperty
      BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628}
      Picture = "Form1.frx":0112
      Key = ""
      EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   Dim nodX As Node
   Dim I As Integer
   Dim J As Integer
   Dim K As Integer
   Dim L As Integer
   
   TreeView1.Nodes.Clear
   '  Corporate Office
   On Error Resume Next
   Set nodX = TreeView1.Nodes.Add(, , "R", "Corporate Office", 1)
   
   '  Regional Level
   For I = 1 To 20
      Set nodX = TreeView1.Nodes.Add("R", tvwChild, "R" & Trim$(Str$(I)), "Region " & Trim$(Str$(I)), 1)
   Next
   nodX.EnsureVisible ' Make sure all nodes are visible.
   
   '  District Level
   J = 1
   For I = 1 To 20
      L = J + 4
      K = J
      For J = K To L
         Set nodX = TreeView1.Nodes.Add("R" & Trim$(Str$(I)), tvwChild, "D" & Trim$(Str$(J)), "District " & Trim$(Str$(J)), 1)
      Next
   Next
   
   '  Store Level
   J = 1
   For I = 1 To 100
      L = J + 9
      K = J
      For J = K To L
         Set nodX = TreeView1.Nodes.Add("D" & Trim$(Str$(I)), tvwChild, "S" & Trim$(Str$(J)), "Store# " & Trim$(Str$(J)), 1)
      Next
   Next
   
End Sub
