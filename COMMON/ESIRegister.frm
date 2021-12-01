VERSION 5.00
Begin VB.Form ESIRegister 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register Fusion"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6270
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5955
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   360
      ScaleHeight     =   1275
      ScaleWidth      =   4515
      TabIndex        =   19
      Top             =   0
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   6015
      Begin VB.Label lblRegOk 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   28
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblProductKey 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   27
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "Product Key:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblPOMOnline 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5160
         TabIndex        =   25
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblFusionOnline 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5160
         TabIndex        =   24
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "POM Users Currently Online:"
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Fusion Users Currently Online:"
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "POM Licenses:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblPOMLicenses 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblExpDate 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblLicenses 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Expiration Date:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Fusion Licenses:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Registration ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Company:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblRegId 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblCompanyName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox txtProdKey 
      Height          =   285
      Index           =   3
      Left            =   4440
      MaxLength       =   5
      TabIndex        =   4
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtProdKey 
      Height          =   285
      Index           =   2
      Left            =   3360
      MaxLength       =   5
      TabIndex        =   3
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtProdKey 
      Height          =   285
      Index           =   1
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   2
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&Register"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtProdKey 
      Height          =   285
      Index           =   0
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   0
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "__"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   18
      Top             =   4275
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "__"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   17
      Top             =   4275
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "__"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   16
      Top             =   4275
      Width           =   255
   End
   Begin VB.Label lblInvalidKey 
      Alignment       =   2  'Center
      Caption         =   "Invalid Product Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   6255
   End
End
Attribute VB_Name = "ESIRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRegister_Click()
    Dim sNewProdKey As String
    sNewProdKey = txtProdKey(0) & "-" & txtProdKey(1) & "-" & txtProdKey(2) & "-" & txtProdKey(3)
    If Len(sNewProdKey) <> 23 And Len(sNewProdKey) <> 24 Then
        lblInvalidKey.Visible = True
        Exit Sub
    End If
    ' Don't forget to fix the text file as well
    If Not NewProductKeyOk(sNewProdKey) Then
        lblInvalidKey.Visible = True
        Exit Sub
    Else
        lblInvalidKey.Visible = False
    End If
    
    If RegisterNewKey(sNewProdKey) Then MsgBox "Registration Successful" Else MsgBox "Registration Failed"
    DisplayRegistrationInfo
End Sub


Private Sub Form_Load()
    Dim sTemp As String
    
    If RegistrationOk(sTemp) Then lblRegOk = "Registration is Valid" Else lblRegOk = "Invalid Registration"
    DisplayRegistrationInfo
End Sub


Sub DisplayRegistrationInfo()
    Dim lRegId As Long
    Dim iUsersAllowed As Integer
    Dim iPOMUsersAllowed As Integer
    Dim dExpirationDte As Date
    Dim iFusionOnline, iPOMOnline As Integer

    'If Not (InStr(1, UCase(Command), "FUSIONROCKS") > 0) Then
        lblRegId = LTrim(Str(RegistrationID))
        lblCompanyName = GetRegCompanyName
        iUsersAllowed = LicensedFusionUsersAllowed
        lblLicenses = LTrim(Str(iUsersAllowed))
        iPOMUsersAllowed = LicensedPOMUsersAllowed
        lblPOMLicenses = LTrim(Str(iPOMUsersAllowed))
        dExpirationDte = GetExpirationDate
        lblExpDate = Format(dExpirationDte, "mm/dd/yyyy")
        lblProductKey = ProductKey
    'Else
     '   lblRegId = "999999"
     '   lblCompanyName = "Key Methods"
     '   lblProductKey = ProductKey
     '   lblLicenses = "Unlimited"
     '   lblPOMLicenses = "Unlimited"
     '   lblExpDate = "<NEVER>"
    'End If
    iFusionOnline = Registration.FusionUsersLoggedIn
    lblFusionOnline = LTrim(Str(iFusionOnline))
    lblPOMOnline = LTrim(Str(iPOMOnline))
    iPOMOnline = Registration.POMUsersLoggedIn
End Sub



Private Sub txtProdKey_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (KeyCode >= 65 And KeyCode <= 90) Or (KeyCode >= 48 And KeyCode <= 57) Or (KeyCode >= 96 And KeyCode <= 105) Then
        If Len(txtProdKey(Index)) = 4 Then
            If Index < 3 Then txtProdKey(Index + 1).SetFocus Else cmdRegister.SetFocus
        End If
    End If
End Sub



