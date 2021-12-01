VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaFavor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Favorites"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   Icon            =   "diaFavor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstFav 
      ForeColor       =   &H00800000&
      Height          =   2400
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   480
      TabIndex        =   2
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
      Top             =   3840
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   300
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   330
      _Version        =   65536
      _ExtentX        =   582
      _ExtentY        =   529
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaFavor.frx":030A
      PictureDn       =   "diaFavor.frx":08AC
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
      FormDesignHeight=   3765
      FormDesignWidth =   3885
   End
   Begin Threed.SSCommand cmdDn 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2760
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaFavor.frx":0E4E
   End
   Begin Threed.SSCommand cmdUp 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2280
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaFavor.frx":1350
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
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   120
      Picture         =   "diaFavor.frx":1852
      Top             =   2880
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   120
      Picture         =   "diaFavor.frx":1D44
      Top             =   2520
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   120
      Picture         =   "diaFavor.frx":2236
      Top             =   1800
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   120
      Picture         =   "diaFavor.frx":2728
      Top             =   2160
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "diaFavor"
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
         sMsg = "No Form Selected Has Been Selected " & vbCr _
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
      cmdDel.enabled = False
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

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs924.htm"
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
         MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Visible = True
         MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Caption = sFavorites(iList)
      Else
         MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Visible = False
         MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Caption = ""
      End If
   Next
   If sFavorites(iList) <> "" Then
      MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Visible = True
      MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Caption = sFavorites(iList)
   Else
      MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Visible = False
      MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(iList))).Caption = ""
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
   Set diaFavor = Nothing
   
End Sub












Private Sub lstFav_Click()
   iIndex = lstFav.ListIndex
   cmdDel.enabled = True
   If iIndex = 0 Then
      cmdUp.Picture = Dsup
      cmdUp.enabled = False
   Else
      cmdUp.Picture = Enup
      cmdUp.enabled = True
   End If
   If iIndex = lstFav.ListCount - 1 Then
      cmdDn.Picture = Dsdn
      cmdDn.enabled = False
   Else
      cmdDn.Picture = Endn
      cmdDn.enabled = True
   End If
   
End Sub
