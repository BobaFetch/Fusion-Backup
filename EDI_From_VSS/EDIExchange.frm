VERSION 5.00
Begin VB.Form EDIMain 
   Caption         =   "EDIMain"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   735
      Left            =   4080
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cmdInvOut 
      Caption         =   "Create EDI Invoice file"
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton cmdASNOut 
      Caption         =   "Create EDI ASN file"
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton cmdImpEDI 
      Caption         =   "Import EDI 830 && 862 data"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "EDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdASNOut_Click()
   EDIAsnOut.Show Modal
End Sub

Private Sub cmdImpEDI_Click()
   EDIImpSO.Show Modal
End Sub

Private Sub cmdInvOut_Click()
   EDIInvOut.Show Modal
End Sub

Private Sub Command1_Click()

Dim strFilePath As String
Dim strFileName As String
Dim bret As Boolean

strFilePath = "C:\Development\FusionCode\EDIFiles\91211\"
'strFileName = "in862.edi(09-09-2011-5.11.01.21A).edi"
strFileName = "in830.edi(09-10-2011-5.11.03.24A).edi"
'ImpEDISalesOrder strFilePath, strFileName

'CreateASNOut
'CreateInvoiceEDIFile

End Sub

Public Function Test()

End Function

