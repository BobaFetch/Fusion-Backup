VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form Recent
   BorderStyle = 1 'Fixed Single
   Caption = "Recent     F6"
   ClientHeight = 3075
   ClientLeft = 990
   ClientTop = 3720
   ClientWidth = 2910
   Icon = "Recent.frx":0000
   KeyPreview = -1 'True
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   ScaleHeight = 3075
   ScaleWidth = 2910
   ShowInTaskbar = 0 'False
   WindowState = 1 'Minimized
   Begin VB.CommandButton cmdCan
      Caption = "&Hide"
      Height = 375
      Left = 1920
      TabIndex = 1
      TabStop = 0 'False
      ToolTipText = "Hides Recent Until Reselected (F6)"
      Top = 2640
      Width = 855
   End
   Begin VB.ListBox lstRec
      Height = 2595
      Left = 120
      TabIndex = 0
      ToolTipText = "Recent List For This Session"
      Top = 0
      Width = 2655
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 120
      Top = 2640
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3075
      FormDesignWidth = 2910
   End
End
Attribute VB_Name = "Recent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Private Sub cmdCan_Click()
   iHideRecent = 1
   SaveSetting "Esi2000", "Programs", "HideRecent", 1
   Unload Me
   
End Sub

Private Sub Form_Activate()
   Refreshrecent
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Hide
   
End Sub


Private Sub Form_Load()
   SetFormSize Me
   Move 0, MdiSect.Height - (Height + 1150)
   Show
   
End Sub


Public Sub Refreshrecent()
   Dim i As Integer
   Dim DataValue As Variant
   Dim DataValues As New Collection
   lstRec.Clear
   For i = 0 To 30
      If Len(Trim(sSession(i))) > 3 Then
         lstRec.AddItem sSession(i)
      Else
         Exit For
      End If
   Next
   On Error Resume Next
   For i = 0 To lstRec.ListCount - 1
      If Len(Trim$(lstRec.List(i))) > 0 Then
         DataValues.Add Item: = Trim(lstRec.List(i)), Key: = Trim(lstRec.List(i))
      End If
   Next
   i = -1
   Erase sSession
   For Each DataValue In DataValues
      i = i + 1
      sSession(i) = DataValue
   Next
   lstRec.Clear
   For i = 0 To 30
      If sSession(i) <> "" Then
         lstRec.AddItem sSession(i)
      Else
         Exit For
      End If
   Next
   If lstRec.ListCount > 0 Then lstRec.ListIndex = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Resize()
   If WindowState = 0 Then Refresh
   
End Sub

Private Sub lstRec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'sSelected = lstRec.List(lstRec.ListIndex)
      'If Len(sSelected) > 0 Then OpenFavorite sSelected
      'WindowState = 1
   End If
   
End Sub


Private Sub lstRec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'sSelected = lstRec.List(lstRec.ListIndex)
   'If Len(sSelected) > 0 Then OpenFavorite sSelected
   'WindowState = 1
   
End Sub
