VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form BompBMe02c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parts List Effectivity"
   ClientHeight    =   5625
   ClientLeft      =   1845
   ClientTop       =   540
   ClientWidth     =   6390
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "BompBMe02c.frx":0000
      Height          =   372
      Left            =   5950
      Picture         =   "BompBMe02c.frx":04F2
      Style           =   1  'Graphical
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   5160
      Width           =   400
   End
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "BompBMe02c.frx":09E4
      Height          =   372
      Left            =   5950
      Picture         =   "BompBMe02c.frx":0ED6
      Style           =   1  'Graphical
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   4780
      Width           =   400
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMe02c.frx":13C8
      Style           =   1  'Graphical
      TabIndex        =   64
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optFrom 
      Caption         =   "From Tree"
      Height          =   255
      Left            =   3360
      TabIndex        =   63
      Top             =   120
      Width           =   1455
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1800
      Top             =   4920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5625
      FormDesignWidth =   6390
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5400
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "Update and Apply Changes"
      Top             =   600
      Width           =   915
   End
   Begin VB.TextBox txtObs 
      Height          =   285
      Index           =   9
      Left            =   4920
      TabIndex        =   44
      Top             =   4440
      Width           =   915
   End
   Begin VB.TextBox txtEff 
      Height          =   285
      Index           =   9
      Left            =   3960
      TabIndex        =   43
      Top             =   4440
      Width           =   915
   End
   Begin VB.TextBox txtObs 
      Height          =   285
      Index           =   8
      Left            =   4920
      TabIndex        =   40
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox txtEff 
      Height          =   285
      Index           =   8
      Left            =   3960
      TabIndex        =   39
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox txtObs 
      Height          =   285
      Index           =   7
      Left            =   4920
      TabIndex        =   36
      Top             =   3720
      Width           =   915
   End
   Begin VB.TextBox txtEff 
      Height          =   285
      Index           =   7
      Left            =   3960
      TabIndex        =   35
      Top             =   3720
      Width           =   915
   End
   Begin VB.TextBox txtObs 
      Height          =   285
      Index           =   6
      Left            =   4920
      TabIndex        =   32
      Top             =   3360
      Width           =   915
   End
   Begin VB.TextBox txtEff 
      Height          =   285
      Index           =   6
      Left            =   3960
      TabIndex        =   31
      Top             =   3360
      Width           =   915
   End
   Begin VB.TextBox txtObs 
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   28
      Top             =   3000
      Width           =   915
   End
   Begin VB.TextBox txtEff 
      Height          =   285
      Index           =   5
      Left            =   3960
      TabIndex        =   27
      Top             =   3000
      Width           =   915
   End
   Begin VB.TextBox txtObs 
      Height          =   285
      Index           =   4
      Left            =   4920
      TabIndex        =   24
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox txtEff 
      Height          =   285
      Index           =   4
      Left            =   3960
      TabIndex        =   23
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox txtObs 
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   20
      Top             =   2280
      Width           =   915
   End
   Begin VB.TextBox txtEff 
      Height          =   285
      Index           =   3
      Left            =   3960
      TabIndex        =   19
      Top             =   2280
      Width           =   915
   End
   Begin VB.TextBox txtObs 
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   16
      Top             =   1920
      Width           =   915
   End
   Begin VB.TextBox txtEff 
      Height          =   285
      Index           =   2
      Left            =   3960
      TabIndex        =   15
      Top             =   1920
      Width           =   915
   End
   Begin VB.TextBox txtObs 
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   12
      Top             =   1560
      Width           =   915
   End
   Begin VB.TextBox txtEff 
      Height          =   285
      Index           =   1
      Left            =   3960
      TabIndex        =   11
      Top             =   1560
      Width           =   915
   End
   Begin VB.TextBox txtObs 
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   4
      Top             =   1200
      Width           =   915
   End
   Begin VB.TextBox txtEff 
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   3
      Top             =   1200
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5400
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   240
      Picture         =   "BompBMe02c.frx":1B76
      Top             =   4995
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   720
      Picture         =   "BompBMe02c.frx":2068
      Top             =   4995
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   480
      Picture         =   "BompBMe02c.frx":255A
      Top             =   4995
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   0
      Picture         =   "BompBMe02c.frx":2A4C
      Top             =   4995
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblPrt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   62
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "**** = Missing or Error In Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   61
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Correct Any Date Discrepancies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   60
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   59
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   58
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label lblMax 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   57
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lblPge 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      TabIndex        =   56
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lblErr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   9
      Left            =   5960
      TabIndex        =   55
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label lblErr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   8
      Left            =   5960
      TabIndex        =   54
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label lblErr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   7
      Left            =   5960
      TabIndex        =   53
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblErr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   6
      Left            =   5960
      TabIndex        =   52
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblErr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   5
      Left            =   5960
      TabIndex        =   51
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblErr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   4
      Left            =   5960
      TabIndex        =   50
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblErr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   5960
      TabIndex        =   49
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblErr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   5960
      TabIndex        =   48
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblErr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   5960
      TabIndex        =   47
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblErr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   5960
      TabIndex        =   46
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   3240
      TabIndex        =   42
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label lblPls 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   41
      Top             =   4440
      Width           =   3075
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   3240
      TabIndex        =   38
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblPls 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   37
      Top             =   4080
      Width           =   3075
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   3240
      TabIndex        =   34
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblPls 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   33
      Top             =   3720
      Width           =   3075
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   3240
      TabIndex        =   30
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblPls 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   29
      Top             =   3360
      Width           =   3075
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   3240
      TabIndex        =   26
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblPls 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   25
      Top             =   3000
      Width           =   3075
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   3240
      TabIndex        =   22
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblPls 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   21
      Top             =   2640
      Width           =   3075
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   18
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblPls 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   3075
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   14
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblPls 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   3075
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   10
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblPls 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Obsolete      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Effective     "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev       "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblPls 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3075
   End
End
Attribute VB_Name = "BompBMe02c"
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
Dim bErrors As Byte

Dim iIndex As Integer
Dim iCurrPage As Integer
Dim iMaxPages As Integer
Dim iTotalLists As Integer

Dim sPartNumber As String

Dim sHeaders(300, 4) As String
'0 = Revision
'1 = Effective
'2 = Obsolete
'3 = Error


Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdDn_Click()
   iCurrPage = iCurrPage + 1
   If iCurrPage > iMaxPages Then iCurrPage = iMaxPages
   GetThisGroup
   
End Sub

Private Sub cmdUp_Click()
   iCurrPage = iCurrPage - 1
   If iCurrPage < 1 Then iCurrPage = 1
   GetThisGroup
   
End Sub


Private Sub cmdUpd_Click()
   Dim A As Integer
   Dim iList As Integer
   Dim bResponse As Byte
   Dim sMsg As String
   A = 0
   CheckBmhDates
   If bErrors Then
      sMsg = "The List Contains Date Errors." & vbCrLf _
             & "Update Parts Lists Anyway?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      bErrors = False
   Else
      bResponse = vbYes
   End If
   If bResponse = vbYes Then
      For iList = 0 To iTotalLists
         MouseCursor 13
         If Len(sHeaders(iList, 1)) > 0 And Len(sHeaders(iList, 2)) > 0 Then
            sSql = "UPDATE BmhdTable SET BMHEFFECTIVE='" & sHeaders(iList, 1) _
                   & "',BMHOBSOLETE='" & sHeaders(iList, 2) _
                   & "' WHERE BMHREF='" & sPartNumber & "' " _
                   & "AND BMHREV='" & sHeaders(iList, 0) & "' "
         Else
            sSql = "UPDATE BmhdTable SET BMHEFFECTIVE=Null" _
                   & ",BMHOBSOLETE=Null" _
                   & " WHERE BMHREF='" & sPartNumber & "' " _
                   & "AND BMHREV='" & sHeaders(iList, 0) & "' "
         End If
         clsADOCon.ExecuteSQL sSql ' rdExecDirect
         If clsADOCon.RowsAffected > 0 Then
            A = A + 1
            If sHeaders(iList, 0) = BompBMe02a.cmbRev Then
               BompBMe02a.txtEff = sHeaders(iList, 1)
               BompBMe02a.txtObs = sHeaders(iList, 2)
            End If
         End If
      Next
      MouseCursor 0
      If A < iTotalLists Then
         MsgBox "Not All Rows Were Updated.", vbInformation, Caption
      Else
         SysMsg str(A) & " Rows Were Updated.", True, Me
      End If
      GetLists
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim iList As Integer
   If MDISect.SideBar.Visible Then
      Move MDISect.Left + 2200, MDISect.Top + 800
   Else
      Move MDISect.Left + 300, MDISect.Top + 800
   End If
   lblPrt = BompBMe02a.cmbPls
   sPartNumber = Compress(lblPrt)
   cmdUp.Picture = Enup
   cmdUp.Enabled = True
   cmdDn.Picture = Endn
   cmdDn.Enabled = True
   For iList = 0 To 9
      lblPls(iList) = lblPrt
   Next
   GetLists
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   If optFrom.value = vbChecked Then
      BompBMe01a.SetFocus
   Else
      BompBMe02a.txtRef.SetFocus
   End If
   Set BompBMe02c = Nothing
   
End Sub



Private Sub GetLists()
   Dim iList As Integer
   iList = -1
   Erase sHeaders
   Dim RdoDte As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREF,BMHPARTNO,BMHREV," _
          & "BMHEFFECTIVE,BMHOBSOLETE FROM BmhdTable " _
          & "WHERE BMHREF='" & sPartNumber & "' " _
          & "ORDER BY BMHOBSOLETE"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte)
   If bSqlRows Then
      With RdoDte
         Do Until .EOF
            iList = iList + 1
            If iList > 299 Then
               sSql = "You Have Too Many Bills For This Part." & vbCrLf _
                      & "You Should Delete Some Old Ones."
               MsgBox sSql, vbInformation, Caption
               Exit Do
            End If
            sHeaders(iList, 0) = "" & Trim(!BMHREV)
            sHeaders(iList, 1) = "" & Format(!BMHEFFECTIVE, "mm/dd/yy")
            sHeaders(iList, 2) = "" & Format(!BMHOBSOLETE, "mm/dd/yy")
            .MoveNext
         Loop
         ClearResultSet RdoDte
      End With
      iTotalLists = iList
      iMaxPages = 1 + (iList \ 10)
      If iMaxPages < 1 Then iMaxPages = 1
      lblPge = " 1"
      lblMax = str(iMaxPages)
      iCurrPage = 1
   Else
      MsgBox "No Parts Lists Found For " & lblPrt & ".", vbExclamation, Caption
      iCurrPage = 0
   End If
   If iMaxPages < 2 Then
      cmdUp.Picture = Dsup
      cmdUp.Enabled = False
      cmdDn.Picture = Dsdn
      cmdDn.Enabled = False
   End If
   Set RdoDte = Nothing
   GetThisGroup
   Exit Sub
   
DiaErr1:
   sProcName = "getlists"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetThisGroup()
   Dim iList As Integer
   Dim A As Integer
   lblPge = str(iCurrPage)
   CheckBmhDates
   On Error GoTo DiaErr1
   A = (iCurrPage - 1) * 10
   iIndex = A
   
   For iList = 0 To 9
      lblPls(iList).Visible = False
      lblRev(iList).Visible = False
      txtEff(iList).Visible = False
      txtObs(iList).Visible = False
      lblErr(iList).Visible = False
   Next
   
   If iCurrPage = 0 Then Exit Sub
   
   For iList = A To A + 9
      If iList > iTotalLists Then Exit For
      lblPls(iList - A).Visible = True
      lblRev(iList - A).Visible = True
      txtEff(iList - A).Visible = True
      txtObs(iList - A).Visible = True
      lblErr(iList - A).Visible = True
      lblRev(iList - A) = sHeaders(iList, 0)
      txtEff(iList - A) = sHeaders(iList, 1)
      txtObs(iList - A) = sHeaders(iList, 2)
      lblErr(iList - A) = sHeaders(iList, 3)
   Next
   On Error Resume Next
   txtEff(0).SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "getthisgr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtEff_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtEff_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtEff_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyDate KeyAscii
   
End Sub

Private Sub txtEff_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub txtEff_LostFocus(Index As Integer)
   If Trim(txtEff(Index)) = "" Then
      txtEff(Index) = CheckDate(txtEff(Index))
   Else
      txtEff(Index) = sHeaders(iIndex + Index, 1)
      Exit Sub
   End If
   sHeaders(iIndex + Index, 1) = "" & txtEff(Index)
   
End Sub

Private Sub txtObs_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtObs_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtObs_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyDate KeyAscii
   
End Sub


Private Sub txtObs_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub



Private Sub CheckBmhDates()
   'There is more to this than shown
   'Need to work with Jerry on the rest
   'See Lost_Focus in txtObs(Index)
   Dim iList As Integer
   Dim A As Integer
   bErrors = False
   For iList = 0 To iTotalLists - 1
      On Error Resume Next
      A = DateDiff("d", sHeaders(iList, 2), sHeaders(iList + 1, 1))
      If A <> 0 Then
         Beep
         sHeaders(iList, 3) = "****"
         bErrors = True
      Else
         sHeaders(iList, 3) = ""
      End If
   Next
   
End Sub

Private Sub txtObs_LostFocus(Index As Integer)
   Dim r As Long
   Dim l As Long
   If Trim(txtObs(Index)) <> "" Then
      txtObs(Index) = CheckDate(txtObs(Index))
   Else
      sHeaders(iIndex + Index, 2) = "" & txtObs(Index)
      Exit Sub
   End If
   On Error Resume Next
   r& = DateValue(txtEff(Index))
   l& = DateValue(txtObs(Index))
   If r& > l& Then
      Beep
      lblErr(Index) = "****"
      txtObs(Index) = txtEff(Index)
   End If
   sHeaders(iIndex + Index, 2) = "" & txtObs(Index)
   
   
End Sub
