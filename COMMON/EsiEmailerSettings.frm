VERSION 5.00
Begin VB.Form EsiEmailerSettings 
   Caption         =   "Fusion Emailer Server Settings"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLogin 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1860
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2880
      TabIndex        =   5
      Top             =   2340
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   1560
      TabIndex        =   4
      Top             =   2340
      Width           =   1035
   End
   Begin VB.TextBox txtDatabaseName 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   1020
      Width           =   2775
   End
   Begin VB.TextBox txtServerName 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Login"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1500
      Width           =   1395
   End
   Begin VB.Label Label4 
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label Label3 
      Caption         =   "To establish email capability on this computer, please enter connection information for the Emailer database:"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   60
      Width           =   4635
   End
   Begin VB.Label Label2 
      Caption         =   "Database Name"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Server Name"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   660
      Width           =   1395
   End
End
Attribute VB_Name = "EsiEmailerSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private emailerServerName As String
Private emailerDatabaseName As String
Private emailerLogin As String
Private emailerPassword As String

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOk_Click()
   SaveSetting "Esi2000", "EsiEmailer", "ServerName", txtServerName.Text
   SaveSetting "Esi2000", "EsiEmailer", "DatabaseName", txtDatabaseName.Text
   SaveSetting "Esi2000", "EsiEmailer", "Login", txtLogin.Text
   SaveSetting "Esi2000", "EsiEmailer", "Password", txtPassword.Text
   Unload Me
End Sub

Private Sub Form_Load()
   emailerServerName = GetSetting("Esi2000", "EsiEmailer", "ServerName", "")
   emailerDatabaseName = GetSetting("Esi2000", "EsiEmailer", "DatabaseName", "")
   emailerLogin = GetSetting("Esi2000", "EsiEmailer", "Login", "")
   emailerPassword = GetSetting("Esi2000", "EsiEmailer", "Password", "")
   txtServerName.Text = emailerServerName
   txtDatabaseName.Text = emailerDatabaseName
   txtLogin.Text = emailerLogin
   txtPassword.Text = emailerPassword
   
End Sub
