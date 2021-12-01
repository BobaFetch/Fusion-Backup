VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LotsLTf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lot Organization - By Part Types"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar statusbar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   13
      Top             =   1785
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkPartType 
      Caption         =   "8"
      Height          =   255
      Index           =   8
      Left            =   5700
      TabIndex        =   11
      Top             =   1320
      Width           =   435
   End
   Begin VB.CheckBox chkPartType 
      Caption         =   "7"
      Height          =   255
      Index           =   7
      Left            =   4950
      TabIndex        =   10
      Top             =   1320
      Width           =   435
   End
   Begin VB.CheckBox chkPartType 
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   4185
      TabIndex        =   9
      Top             =   1320
      Width           =   435
   End
   Begin VB.CheckBox chkPartType 
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   3435
      TabIndex        =   8
      Top             =   1320
      Width           =   435
   End
   Begin VB.CheckBox chkPartType 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   2685
      TabIndex        =   7
      Top             =   1320
      Width           =   435
   End
   Begin VB.CheckBox chkPartType 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   1935
      TabIndex        =   6
      Top             =   1320
      Width           =   435
   End
   Begin VB.CheckBox chkPartType 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   1170
      TabIndex        =   5
      Top             =   1320
      Width           =   435
   End
   Begin VB.CheckBox chkPartType 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   4
      Top             =   1320
      Width           =   435
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTf03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdOrg 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   420
      Left            =   6960
      TabIndex        =   1
      ToolTipText     =   "Reorganize Lots And Inventory"
      Top             =   1140
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   420
      Left            =   6960
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2100
      FormDesignWidth =   8085
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Process all parts of the following types:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   12
      Top             =   1020
      Width           =   6375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   $"LotsLTf03a.frx":07AE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   420
      TabIndex        =   2
      Top             =   240
      Width           =   6315
   End
End
Attribute VB_Name = "LotsLTf03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 6/30/03
Option Explicit

Private Sub chkPartType_Click(Index As Integer)
   Dim I As Integer
   cmdOrg.Enabled = False
   For I = 1 To 8
      If chkPartType(I).value = vbChecked Then
         cmdOrg.Enabled = True
         Exit Sub
      End If
   Next
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 5550
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdOrg_Click()
   
   'determine the number of parts selected
   Dim sWhereIn As String
   Dim I As Long
   Dim Count As Long
   For I = 1 To 8
      If Me.chkPartType(I).value = vbChecked Then
         If sWhereIn = "" Then
            sWhereIn = "where PALEVEL in ( " & I
         Else
            sWhereIn = sWhereIn & ", " & I
         End If
      End If
   Next
   sWhereIn = sWhereIn & " )"
   If sWhereIn = " )" Or InStr(1, sWhereIn, "(  )") > 0 Then
      MsgBox "No part types selected"
      Exit Sub
   End If
   
   Dim rdo As ADODB.Recordset
   'sSql = "select count(*) from PartTable " & sWhereIn
   
   Dim sFrom As String
   sFrom = "From PartTable" & vbCrLf _
           & "Join (Select LOTPARTREF, sum(LOIQUANTITY) as LOTQTY" & vbCrLf _
           & "from LoitTable join LohdTable on LOINUMBER = LOTNUMBER" & vbCrLf _
           & "group by LOTPARTREF) as x" & vbCrLf _
           & "on x.LOTPARTREF = PARTREF" & vbCrLf _
           & "and x.LOTQTY <> PAQOH" & vbCrLf _
           & sWhereIn & vbCrLf
   
   sSql = "select count(*)" & vbCrLf & sFrom
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bSqlRows Then
      Count = rdo.Fields(0)
      If Count = 0 Then
         MsgBox "There are no parts of the selected types with mismatching part quantity and lot quantity."
         Set rdo = Nothing
         Exit Sub
      End If
      Select Case MsgBox("Reorganize " & Count & " parts of types selected?", _
                         vbYesNo + vbQuestion, Caption)
         Case vbYes
            'process parts
         Case Else
            Set rdo = Nothing
            Exit Sub
      End Select
      
   Else
      Set rdo = Nothing
      Exit Sub
   End If
   Set rdo = Nothing
   'process all parts of desired types where lot qty <> paqoh
   MouseCursor ccHourglass
   statusbar.SimpleText = ""
   cmdOrg.Enabled = False
   sSql = "select PARTREF, PAQOH, LOTQTY" & vbCrLf & sFrom & "order by x.LOTPARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   I = 0
   If bSqlRows Then
      Dim lot As New ClassLot
      Do While Not rdo.EOF
         Dim PartNo As String
         Dim qoh As Currency, lotQty As Currency
         Dim userLotID As String
         
         PartNo = Trim(rdo.Fields(0))
         qoh = rdo.Fields(1)
         lotQty = rdo.Fields(2)
         If qoh < 0 Then qoh = 0
         I = I + 1
         statusbar.SimpleText = I & "/" & Count & "   " & PartNo & " qoh = " & qoh
'         Debug.Print statusbar.SimpleText
         userLotID = "REORG-" & PartNo & "-" & Format(ES_SYSDATE, "mm/dd/yy")
         If lot.ConsolidateLots(PartNo, qoh, userLotID) Then
         Else
            Debug.Print "failed"
         End If
         DoEvents
         rdo.MoveNext
      Loop
   End If
   MouseCursor ccArrow
   cmdOrg.Enabled = True
   Set rdo = Nothing
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   MouseCursor ccArrow
End Sub


Private Sub Form_Resize()
   Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set LotsLTf03a = Nothing
End Sub
