VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SysFavorite 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Favorites"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   ForeColor       =   &H8000000F&
   HelpContextID   =   924
   Icon            =   "SysFavorite.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "SysFavorite.frx":030A
      DownPicture     =   "SysFavorite.frx":07FC
      Enabled         =   0   'False
      Height          =   372
      Left            =   3360
      MaskColor       =   &H00000000&
      Picture         =   "SysFavorite.frx":0CEE
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2664
      Width           =   400
   End
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "SysFavorite.frx":11E0
      DownPicture     =   "SysFavorite.frx":16D2
      Enabled         =   0   'False
      Height          =   372
      Left            =   3360
      MaskColor       =   &H00000000&
      Picture         =   "SysFavorite.frx":1BC4
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2280
      Width           =   400
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SysFavorite.frx":20B6
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ListBox lstFav 
      ForeColor       =   &H00800000&
      Height          =   2205
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   1560
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4080
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3360
      Top             =   3240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3750
      FormDesignWidth =   3885
   End
   Begin VB.Label lblDsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Or Remove A Section Favorite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   492
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   2412
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   2880
      Picture         =   "SysFavorite.frx":2864
      Top             =   360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   3360
      Picture         =   "SysFavorite.frx":2D56
      Top             =   360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   3600
      Picture         =   "SysFavorite.frx":3248
      Top             =   360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   3120
      Picture         =   "SysFavorite.frx":373A
      Top             =   360
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "SysFavorite"
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

Dim iIndex As Integer

Private Sub cmdAdd_Click()
   Dim bByte As Byte
   Dim iRow As Integer
   Dim sMsg As String
   
   If lstFav.ListCount < 12 Then
      bByte = False
      If Len(sCurrForm) > 0 Then
         For iRow = 0 To lstFav.ListCount - 1
            If sCurrForm = lstFav.List(iRow) Then bByte = True
         Next
         If Not bByte Then
            lstFav.AddItem sCurrForm
         Else
            'SysSysSysBeep
         End If
      Else
         sMsg = "No Form Selected Has Been Selected " & vbCrLf _
                & "Or The Selected Form Doesn't Accept Favorites."
         MsgBox sMsg, vbInformation, Caption
      End If
   Else
      MsgBox "List Limit Of Favorites is 12.", vbInformation, Caption
   End If
   
End Sub

Private Sub cmdCan_Click()
   Form_Deactivate
   
End Sub

Private Sub cmdDel_Click()
   If iIndex >= 0 Then
      lstFav.RemoveItem lstFav.ListIndex
      cmdDel.Enabled = False
   End If
   
End Sub

Private Sub cmdDn_Click()
   Dim sText As String
   If iIndex < lstFav.ListCount - 1 Then
      sText = lstFav.List(iIndex + 1)
      lstFav.List(iIndex + 1) = lstFav.List(iIndex)
      lstFav.List(iIndex) = sText
      lstFav.Selected(iIndex + 1) = True
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 924
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdUp_Click()
   Dim sText As String
   If iIndex > 0 Then
      sText = lstFav.List(iIndex - 1)
      lstFav.List(iIndex - 1) = lstFav.List(iIndex)
      lstFav.List(iIndex) = sText
      lstFav.Selected(iIndex - 1) = True
   End If
   
End Sub

Private Sub Form_Activate()
   Caption = sProgName & " " & Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   Dim iList As Integer
   Move 500, 500
   For iList = 1 To 11
      If Trim(sFavorites(iList)) <> "" Then lstFav.AddItem sFavorites(iList)
   Next
   If Trim(sFavorites(iList)) <> "" Then lstFav.AddItem sFavorites(iList)
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim iList As Integer
   Dim sSection As String
   Select Case sProgName
      Case "Engineering"
         sSection = "EsiEngr"
      Case "Quality Assurance"
         sSection = "EsiQual"
      Case "Production"
         sSection = "EsiProd"
      Case "Administration"
         sSection = "EsiAdmn"
      Case "Finance"
         sSection = "EsiFina"
      Case "Sales"
         sSection = "EsiSale"
      Case "Inventory"
         sSection = "EsiInvc"
      Case Else
         sSection = ""
   End Select
   Erase sFavorites
   For iList = 1 To lstFav.ListCount
      sFavorites(iList) = "" & Trim(lstFav.List(iList - 1))
   Next
   For iList = 1 To 11
      If sFavorites(iList) <> "" Then
         MDISect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Visible = True
         MDISect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Caption = sFavorites(iList)
      Else
         MDISect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Visible = False
         MDISect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Caption = ""
      End If
   Next
   If sFavorites(iList) <> "" Then
      MDISect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Visible = True
      MDISect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Caption = sFavorites(iList)
   Else
      MDISect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Visible = False
      MDISect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Caption = ""
   End If
   
   If sSection <> "" Then
      For iList = 1 To 9
         SaveSetting "Esi2000", sSection, "Favorite" & Trim(Str(iList)), sFavorites(iList)
      Next
      SaveSetting "Esi2000", sSection, "Favorite" & Trim(Str(iList)), sFavorites(iList)
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set SysFavorite = Nothing
   
End Sub












Private Sub lstFav_Click()
   iIndex = lstFav.ListIndex
   cmdDel.Enabled = True
   If iIndex = 0 Then
      cmdUp.Picture = Dsup
      cmdUp.Enabled = False
   Else
      cmdUp.Picture = Enup
      cmdUp.Enabled = True
   End If
   If iIndex = lstFav.ListCount - 1 Then
      cmdDn.Picture = Dsdn
      cmdDn.Enabled = False
   Else
      cmdDn.Picture = Endn
      cmdDn.Enabled = True
   End If
   
End Sub
