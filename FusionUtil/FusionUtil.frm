VERSION 5.00
Begin VB.Form FusionMainUtil 
   Caption         =   "Fusion Util"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbConfig 
      Caption         =   "Generate Fusion ConfigFile"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "FusionMainUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbConfig_Click()
   GenFusConfigFile.Show Modal
End Sub
