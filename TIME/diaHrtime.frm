VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form diaHrtme 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter/Revise Daily Time Charges"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkDelete 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   0
      Left            =   8160
      TabIndex        =   13
      ToolTipText     =   "Checked For Direct Labor"
      Top             =   2160
      Width           =   360
   End
   Begin VB.CheckBox chkDelete 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   23
      Top             =   2520
      Width           =   360
   End
   Begin VB.CheckBox chkDelete 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   8160
      TabIndex        =   33
      Top             =   2880
      Width           =   360
   End
   Begin VB.CheckBox chkDelete 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   8160
      TabIndex        =   43
      Top             =   3240
      Width           =   360
   End
   Begin VB.CheckBox chkDelete 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   8160
      TabIndex        =   53
      Top             =   3600
      Width           =   360
   End
   Begin VB.CheckBox chkDelete 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   8160
      TabIndex        =   63
      Top             =   3960
      Width           =   360
   End
   Begin VB.CheckBox chkDelete 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   6
      Left            =   8160
      TabIndex        =   73
      Top             =   4320
      Width           =   360
   End
   Begin VB.CheckBox chkDelete 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   7
      Left            =   8160
      TabIndex        =   83
      Top             =   4680
      Width           =   360
   End
   Begin VB.ComboBox cboFind 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   540
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4755
   End
   Begin VB.CheckBox chkElapsed 
      Caption         =   "&A____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1740
      TabIndex        =   128
      TabStop         =   0   'False
      ToolTipText     =   "Sets The Nex Beginning Time (Local Setting)"
      Top             =   1200
      Width           =   850
   End
   Begin VB.TextBox txtHrs 
      Height          =   315
      Index           =   7
      Left            =   6720
      TabIndex        =   82
      ToolTipText     =   "Operation If MO"
      Top             =   4680
      Width           =   540
   End
   Begin VB.TextBox txtHrs 
      Height          =   315
      Index           =   6
      Left            =   6720
      TabIndex        =   72
      ToolTipText     =   "Operation If MO"
      Top             =   4320
      Width           =   540
   End
   Begin VB.TextBox txtHrs 
      Height          =   315
      Index           =   5
      Left            =   6720
      TabIndex        =   62
      ToolTipText     =   "Operation If MO"
      Top             =   3960
      Width           =   540
   End
   Begin VB.TextBox txtHrs 
      Height          =   315
      Index           =   4
      Left            =   6720
      TabIndex        =   52
      ToolTipText     =   "Operation If MO"
      Top             =   3600
      Width           =   540
   End
   Begin VB.TextBox txtHrs 
      Height          =   315
      Index           =   3
      Left            =   6720
      TabIndex        =   42
      ToolTipText     =   "Operation If MO"
      Top             =   3240
      Width           =   540
   End
   Begin VB.TextBox txtHrs 
      Height          =   315
      Index           =   2
      Left            =   6720
      TabIndex        =   32
      ToolTipText     =   "Operation If MO"
      Top             =   2880
      Width           =   540
   End
   Begin VB.TextBox txtHrs 
      Height          =   315
      Index           =   1
      Left            =   6720
      TabIndex        =   22
      ToolTipText     =   "Operation If MO"
      Top             =   2520
      Width           =   540
   End
   Begin VB.TextBox txtHrs 
      Height          =   315
      Index           =   0
      Left            =   6720
      TabIndex        =   12
      ToolTipText     =   "Operation If MO"
      Top             =   2160
      Width           =   540
   End
   Begin VB.CommandButton cmdOps 
      DownPicture     =   "diaHrtime.frx":0000
      Height          =   315
      Left            =   5400
      Picture         =   "diaHrtime.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   125
      TabStop         =   0   'False
      ToolTipText     =   "MO Completions"
      Top             =   1560
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CheckBox optNew 
      Caption         =   "New Card"
      Height          =   255
      Left            =   360
      TabIndex        =   124
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkAuto 
      Caption         =   "&A____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1740
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Sets The Nex Beginning Time (Local Setting)"
      Top             =   960
      Width           =   850
   End
   Begin VB.ListBox lstItm 
      Height          =   255
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   122
      Top             =   5490
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CheckBox optInd 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   960
      Width           =   850
   End
   Begin ComctlLib.ProgressBar prg1 
      Height          =   255
      Left            =   1200
      TabIndex        =   120
      Top             =   5160
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox txtSri 
      Height          =   315
      Index           =   7
      Left            =   4400
      TabIndex        =   78
      Tag             =   "3"
      ToolTipText     =   "Setup, Run, Tooling or Indirect"
      Top             =   4680
      Width           =   360
   End
   Begin VB.TextBox txtSri 
      Height          =   315
      Index           =   6
      Left            =   4400
      TabIndex        =   68
      Tag             =   "3"
      ToolTipText     =   "Setup, Run, Tooling or Indirect"
      Top             =   4320
      Width           =   360
   End
   Begin VB.TextBox txtSri 
      Height          =   315
      Index           =   5
      Left            =   4400
      TabIndex        =   58
      Tag             =   "3"
      ToolTipText     =   "Setup, Run, Tooling or Indirect"
      Top             =   3960
      Width           =   360
   End
   Begin VB.TextBox txtSri 
      Height          =   315
      Index           =   4
      Left            =   4400
      TabIndex        =   48
      Tag             =   "3"
      ToolTipText     =   "Setup, Run, Tooling or Indirect"
      Top             =   3600
      Width           =   360
   End
   Begin VB.TextBox txtSri 
      Height          =   315
      Index           =   3
      Left            =   4400
      TabIndex        =   38
      Tag             =   "3"
      ToolTipText     =   "Setup, Run, Tooling or Indirect"
      Top             =   3240
      Width           =   360
   End
   Begin VB.TextBox txtSri 
      Height          =   315
      Index           =   2
      Left            =   4400
      TabIndex        =   28
      Tag             =   "3"
      ToolTipText     =   "Setup, Run, Tooling or Indirect"
      Top             =   2880
      Width           =   360
   End
   Begin VB.TextBox txtSri 
      Height          =   315
      Index           =   1
      Left            =   4400
      TabIndex        =   18
      Tag             =   "3"
      ToolTipText     =   "Setup, Run, Tooling or Indirect"
      Top             =   2520
      Width           =   360
   End
   Begin VB.TextBox txtSri 
      Height          =   315
      Index           =   0
      Left            =   4380
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Setup, Run, Tooling or Indirect"
      Top             =   2160
      Width           =   360
   End
   Begin VB.CommandButton cmdFind 
      DownPicture     =   "diaHrtime.frx":09B4
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      Picture         =   "diaHrtime.frx":0CF6
      Style           =   1  'Graphical
      TabIndex        =   116
      TabStop         =   0   'False
      ToolTipText     =   "Find A Manufacturing Order"
      Top             =   1560
      Width           =   350
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "diaHrtime.frx":1038
      Height          =   315
      Left            =   5760
      Picture         =   "diaHrtime.frx":1512
      Style           =   1  'Graphical
      TabIndex        =   115
      TabStop         =   0   'False
      ToolTipText     =   "Show Time Card"
      Top             =   1560
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "C&ancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6180
      TabIndex        =   114
      TabStop         =   0   'False
      ToolTipText     =   "Cancel Time Card Entry"
      Top             =   1560
      Width           =   875
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   74
      Top             =   4680
      Value           =   1  'Checked
      Width           =   360
   End
   Begin VB.TextBox txtMon 
      Height          =   315
      Index           =   7
      Left            =   500
      TabIndex        =   75
      Top             =   4680
      Width           =   2655
   End
   Begin VB.TextBox txtRun 
      Height          =   315
      Index           =   7
      Left            =   3220
      TabIndex        =   76
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox txtOpn 
      Height          =   315
      Index           =   7
      Left            =   3880
      TabIndex        =   77
      Top             =   4680
      Width           =   480
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   7
      Left            =   4780
      Style           =   2  'Dropdown List
      TabIndex        =   79
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   4680
      Width           =   660
   End
   Begin VB.TextBox txtEnd 
      Height          =   315
      Index           =   7
      Left            =   6080
      TabIndex        =   81
      Tag             =   "5"
      Text            =   "  :"
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox txtBeg 
      Height          =   315
      Index           =   7
      Left            =   5440
      TabIndex        =   80
      Tag             =   "5"
      Text            =   "  :"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   64
      Top             =   4320
      Value           =   1  'Checked
      Width           =   360
   End
   Begin VB.TextBox txtMon 
      Height          =   315
      Index           =   6
      Left            =   500
      TabIndex        =   65
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox txtRun 
      Height          =   315
      Index           =   6
      Left            =   3220
      TabIndex        =   66
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtOpn 
      Height          =   315
      Index           =   6
      Left            =   3880
      TabIndex        =   67
      Top             =   4320
      Width           =   480
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   6
      Left            =   4780
      Style           =   2  'Dropdown List
      TabIndex        =   69
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   4320
      Width           =   660
   End
   Begin VB.TextBox txtEnd 
      Height          =   315
      Index           =   6
      Left            =   6080
      TabIndex        =   71
      Tag             =   "5"
      Text            =   "  :"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtBeg 
      Height          =   315
      Index           =   6
      Left            =   5440
      TabIndex        =   70
      Tag             =   "5"
      Text            =   "  :"
      Top             =   4320
      Width           =   615
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   54
      Top             =   3960
      Value           =   1  'Checked
      Width           =   360
   End
   Begin VB.TextBox txtMon 
      Height          =   315
      Index           =   5
      Left            =   500
      TabIndex        =   55
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox txtRun 
      Height          =   315
      Index           =   5
      Left            =   3220
      TabIndex        =   56
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox txtOpn 
      Height          =   315
      Index           =   5
      Left            =   3880
      TabIndex        =   57
      Top             =   3960
      Width           =   480
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   5
      Left            =   4780
      Style           =   2  'Dropdown List
      TabIndex        =   59
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   3960
      Width           =   660
   End
   Begin VB.TextBox txtEnd 
      Height          =   315
      Index           =   5
      Left            =   6080
      TabIndex        =   61
      Tag             =   "5"
      Text            =   "  :"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox txtBeg 
      Height          =   315
      Index           =   5
      Left            =   5440
      TabIndex        =   60
      Tag             =   "5"
      Text            =   "  :"
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   44
      Top             =   3600
      Value           =   1  'Checked
      Width           =   360
   End
   Begin VB.TextBox txtMon 
      Height          =   315
      Index           =   4
      Left            =   500
      TabIndex        =   45
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtRun 
      Height          =   315
      Index           =   4
      Left            =   3220
      TabIndex        =   46
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtOpn 
      Height          =   315
      Index           =   4
      Left            =   3880
      TabIndex        =   47
      Top             =   3600
      Width           =   480
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   4
      Left            =   4780
      Style           =   2  'Dropdown List
      TabIndex        =   49
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   3600
      Width           =   660
   End
   Begin VB.TextBox txtEnd 
      Height          =   315
      Index           =   4
      Left            =   6080
      TabIndex        =   51
      Tag             =   "5"
      Text            =   "  :"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtBeg 
      Height          =   315
      Index           =   4
      Left            =   5440
      TabIndex        =   50
      Tag             =   "5"
      Text            =   "  :"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtEnd 
      Height          =   315
      Index           =   2
      Left            =   6080
      TabIndex        =   31
      Tag             =   "5"
      Text            =   "  :"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   34
      Top             =   3240
      Value           =   1  'Checked
      Width           =   360
   End
   Begin VB.TextBox txtMon 
      Height          =   315
      Index           =   3
      Left            =   500
      TabIndex        =   35
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtRun 
      Height          =   315
      Index           =   3
      Left            =   3220
      TabIndex        =   36
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox txtOpn 
      Height          =   315
      Index           =   3
      Left            =   3880
      TabIndex        =   37
      Top             =   3240
      Width           =   480
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   4780
      Style           =   2  'Dropdown List
      TabIndex        =   39
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   3240
      Width           =   660
   End
   Begin VB.TextBox txtEnd 
      Height          =   315
      Index           =   3
      Left            =   6080
      TabIndex        =   41
      Tag             =   "5"
      Text            =   "  :"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox txtBeg 
      Height          =   315
      Index           =   3
      Left            =   5440
      TabIndex        =   40
      Tag             =   "5"
      Text            =   "  :"
      Top             =   3240
      Width           =   615
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   2880
      Value           =   1  'Checked
      Width           =   360
   End
   Begin VB.TextBox txtMon 
      Height          =   315
      Index           =   2
      Left            =   500
      TabIndex        =   25
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtRun 
      Height          =   315
      Index           =   2
      Left            =   3220
      TabIndex        =   26
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtOpn 
      Height          =   315
      Index           =   2
      Left            =   3880
      TabIndex        =   27
      Top             =   2880
      Width           =   480
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   4780
      Style           =   2  'Dropdown List
      TabIndex        =   29
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   2880
      Width           =   660
   End
   Begin VB.TextBox txtBeg 
      Height          =   315
      Index           =   2
      Left            =   5440
      TabIndex        =   30
      Tag             =   "5"
      Text            =   "  :"
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtBeg 
      Height          =   315
      Index           =   1
      Left            =   5440
      TabIndex        =   20
      Tag             =   "5"
      Text            =   "  :"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox txtEnd 
      Height          =   315
      Index           =   1
      Left            =   6080
      TabIndex        =   21
      Tag             =   "5"
      Text            =   "  :"
      Top             =   2520
      Width           =   615
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   4780
      Style           =   2  'Dropdown List
      TabIndex        =   19
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   2520
      Width           =   660
   End
   Begin VB.TextBox txtOpn 
      Height          =   315
      Index           =   1
      Left            =   3880
      TabIndex        =   17
      Top             =   2520
      Width           =   480
   End
   Begin VB.TextBox txtRun 
      Height          =   315
      Index           =   1
      Left            =   3220
      TabIndex        =   16
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox txtMon 
      Height          =   315
      Index           =   1
      Left            =   500
      TabIndex        =   15
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Value           =   1  'Checked
      Width           =   360
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7080
      TabIndex        =   102
      TabStop         =   0   'False
      ToolTipText     =   "Enter Updated Time Card"
      Top             =   1560
      Width           =   875
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Checked For Direct Labor"
      Top             =   2160
      Value           =   1  'Checked
      Width           =   360
   End
   Begin VB.TextBox txtMon 
      Height          =   315
      Index           =   0
      Left            =   500
      TabIndex        =   5
      ToolTipText     =   "Manufacturing Order (D), Account For (I)"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtRun 
      Height          =   315
      Index           =   0
      Left            =   3220
      TabIndex        =   6
      ToolTipText     =   "Run Number If MO"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtOpn 
      Height          =   315
      Index           =   0
      Left            =   3880
      TabIndex        =   7
      ToolTipText     =   "Operation If MO"
      Top             =   2160
      Width           =   480
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   4780
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   "Select Type Code From List"
      Top             =   2160
      Width           =   660
   End
   Begin VB.TextBox txtEnd 
      Height          =   315
      Index           =   0
      Left            =   6080
      TabIndex        =   11
      Tag             =   "5"
      Text            =   "  :"
      ToolTipText     =   "Ending Time"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtBeg 
      Height          =   315
      Index           =   0
      Left            =   5440
      TabIndex        =   10
      Tag             =   "5"
      Text            =   "  :"
      ToolTipText     =   "Starting Time"
      Top             =   2160
      Width           =   615
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Tag             =   "4"
      Top             =   280
      Width           =   1095
   End
   Begin VB.ComboBox cmbEmp 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select From List Or Enter Number"
      Top             =   280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7080
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   60
      TabIndex        =   85
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaHrtime.frx":19EC
      PictureDn       =   "diaHrtime.frx":1B32
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   5400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5820
      FormDesignWidth =   8520
   End
   Begin Threed.SSCommand cmdDn 
      Height          =   375
      Left            =   7500
      TabIndex        =   110
      TabStop         =   0   'False
      ToolTipText     =   "Next Page (Page Down)"
      Top             =   5400
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaHrtime.frx":1C78
   End
   Begin Threed.SSCommand cmdUp 
      Height          =   375
      Left            =   7500
      TabIndex        =   111
      TabStop         =   0   'False
      ToolTipText     =   "Last Page (Page Up)"
      Top             =   5040
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaHrtime.frx":217A
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Del"
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
      Index           =   19
      Left            =   8160
      TabIndex        =   130
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Elapsed Time"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   127
      ToolTipText     =   "Sets The Nex Beginning Time (Local Setting)"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Error?    "
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
      Index           =   17
      Left            =   7320
      TabIndex        =   126
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Auto Start Time Is On"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   123
      ToolTipText     =   "Sets The Nex Beginning Time (Local Setting)"
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Mark All Indirect"
      Height          =   255
      Index           =   15
      Left            =   2880
      TabIndex        =   121
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "S R I"
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
      Index           =   14
      Left            =   4400
      TabIndex        =   119
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblWen 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   7080
      TabIndex        =   118
      ToolTipText     =   "Week Ending (System Administration Setup)"
      Top             =   615
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Week Ending"
      Height          =   255
      Index           =   12
      Left            =   5880
      TabIndex        =   117
      Top             =   615
      Width           =   1215
   End
   Begin VB.Label lblPge 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6360
      TabIndex        =   113
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   255
      Index           =   13
      Left            =   5640
      TabIndex        =   112
      Top             =   5160
      Width           =   615
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   840
      Picture         =   "diaHrtime.frx":267C
      Top             =   5040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   120
      Picture         =   "diaHrtime.frx":2B6E
      Top             =   5040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   360
      Picture         =   "diaHrtime.frx":3060
      Top             =   5040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   600
      Picture         =   "diaHrtime.frx":3552
      Top             =   5040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblErrors 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   7320
      TabIndex        =   109
      Top             =   4680
      Width           =   780
   End
   Begin VB.Label lblErrors 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   7320
      TabIndex        =   108
      Top             =   4320
      Width           =   780
   End
   Begin VB.Label lblErrors 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   7320
      TabIndex        =   107
      Top             =   3960
      Width           =   780
   End
   Begin VB.Label lblErrors 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   7320
      TabIndex        =   106
      Top             =   3600
      Width           =   780
   End
   Begin VB.Label lblErrors 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   7320
      TabIndex        =   105
      Top             =   3240
      Width           =   780
   End
   Begin VB.Label lblErrors 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   7320
      TabIndex        =   104
      Top             =   2880
      Width           =   780
   End
   Begin VB.Label lblErrors 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   7320
      TabIndex        =   103
      Top             =   2520
      Width           =   780
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   7080
      TabIndex        =   101
      ToolTipText     =   "Today's Accumulated Hours"
      Top             =   960
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Card Total"
      Height          =   255
      Index           =   11
      Left            =   5880
      TabIndex        =   100
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hours       "
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
      Index           =   10
      Left            =   6720
      TabIndex        =   99
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End       "
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
      Index           =   9
      Left            =   6080
      TabIndex        =   98
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start      "
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
      Index           =   8
      Left            =   5540
      TabIndex        =   97
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tp     "
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
      Index           =   7
      Left            =   4780
      TabIndex        =   96
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Op No"
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
      Index           =   6
      Left            =   3880
      TabIndex        =   95
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run      "
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
      Index           =   5
      Left            =   3220
      TabIndex        =   94
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturing Order/Account           "
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
      Index           =   4
      Left            =   500
      TabIndex        =   93
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "D/I"
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
      Left            =   120
      TabIndex        =   92
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblErrors 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   0
      Left            =   7320
      TabIndex        =   91
      ToolTipText     =   "Hours For Entry"
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Card Date"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   90
      Top             =   280
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   89
      Top             =   610
      Width           =   1575
   End
   Begin VB.Label lblSsn 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4680
      TabIndex        =   88
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   87
      Top             =   610
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   86
      Top             =   280
      Width           =   1575
   End
End
Attribute VB_Name = "diaHrtme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/5/02 - Added fill for TMDATE
'10/7/03 - Added Operations dialog
Option Explicit

Private Const CURSOR_HOURGLASS = 13
Private Const CURSOR_NORMAL = 0

Private Const LINES_PER_PAGE = 8
'Private Const MAX_PAGES = 8  no longer used


Dim iTotalPages As Integer


'Dim RdoQry1 As rdoQuery
'Dim RdoQry2 As rdoQuery

Dim cmdObj1 As ADODB.Command
Dim cmdObj2 As ADODB.Command

Dim prmObj1 As ADODB.Parameter
Dim prmObj2 As ADODB.Parameter
Dim prmObj3 As ADODB.Parameter

Dim bAddingCard As Byte
Dim bCanceled As Byte
Dim bOnLoad As Byte
Dim bGoodCard As Byte
Dim bGoodEmployee As Byte
Dim bGoodTime As Boolean

Dim iIndex As Integer
Dim iCurrIndex As Integer
Dim iCurrPage As Integer
Dim iTotalCenters As Integer
Dim iTotalCodes As Integer
Dim iTotalEntries As Integer

Dim cEmplRate As Currency
Dim cShopOhRate As Currency
Dim cShopOhFixed As Currency

Dim sNewCard As String
Dim sOldCard As String
Dim sEmplCenter As String
Dim sEmplShop As String
Dim sEmplAcct As String
Dim sCardNumber As String

'Company Accounts
Dim sCoTimeAcct As String
Dim sCoLaborAcct As String
Dim sWorkCenterAcct As String

'timecard arrays
Private vTypeCode(100, 2) As Variant
Private vWorkCenter(200, 3) As Variant
Private Const WC_CODE = 0
Private Const WC_OHPERCENT = 1
Private Const WC_OHFIXED = 2

Private vTimeCard() As Variant
'0 = D/I
'1 = Manufacturing Order
'2 = Acct
'3 = Run
'4 = Op
'5 = Type
'6 = Beg
'7 = End
'8 = Hours
'9 = Setup, Run, Tooling or Indirect
'10 = Workcenter
'11 = Time Account
Private Const TC_DI = 0    '= 1 for direct (use TC_SRI <> "I" instead
Private Const TC_MO = 1
Private Const TC_ACCT = 2 'indirect time charge account
Private Const TC_RUN = 3
Private Const TC_OP = 4
Private Const TC_TYPE = 5
Private Const TC_BEGIN = 6
Private Const TC_END = 7
Private Const TC_HOURS = 8
Private Const TC_SRI = 9      'S = Setup, R = Run, I = indirect
Private Const TC_WC = 10
Private Const TC_TIMEACCT = 11
Private Const TC_DELETE = 12
Private Const TC_ACCEPTED = 13
Private Const TC_REJECTED = 14
Private Const TC_SCRAPPED = 15
Private Const TC_STARTTIME = 16
Private Const TC_ENDTIME = 17
Private Const TC_COMMENTS = 18

Private bGood() As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetShopInfo()
   On Error GoTo DiaErr1
   Dim RdoShop As ADODB.Recordset
   sSql = "SELECT SHPREF,SHPACCT,SHPOHRATE,SHPOHTOTAL FROM ShopTable " _
          & "WHERE SHPREF='" & sEmplShop & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShop, ES_FORWARD)
   If bSqlRows Then
      With RdoShop
         sWorkCenterAcct = "" & Trim(!SHPACCT)
         cShopOhRate = !SHPOHRATE
         cShopOhFixed = !SHPOHTOTAL
         .Cancel
      End With
   Else
      cShopOhRate = 0
      cShopOhFixed = 0
      sWorkCenterAcct = ""
   End If
   'If cShopOhRate = 0 Then cShopOhRate = 1
   Set RdoShop = Nothing
   Exit Sub
   
DiaErr1:
   cShopOhRate = 1
   cShopOhFixed = 0
   
End Sub

Private Sub GetWorkCenterInfo()
   On Error GoTo DiaErr1
   Dim RdoShop As ADODB.Recordset
   sSql = "SELECT WCNREF,WCNACCT,WCNOHPCT,WCNOHFIXED FROM WcntTable " _
          & "WHERE WCNREF='" & sEmplCenter & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShop, ES_FORWARD)
   If bSqlRows Then
      With RdoShop
         sWorkCenterAcct = "" & Trim(!WCNACCT)
         cShopOhRate = !WCNOHPCT
         cShopOhFixed = !WCNOHFIXED
         .Cancel
      End With
   Else
      cShopOhRate = 0
      cShopOhFixed = 0
      sWorkCenterAcct = ""
   End If
   'If cShopOhRate = 0 Then cShopOhRate = 1
   Set RdoShop = Nothing
   Exit Sub
   
DiaErr1:
   cShopOhRate = 1
   cShopOhFixed = 0
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   
   'determine whether to use auto start time
   sOptions = GetSetting("Esi2000", "EsiAdmn", "AutoTime", sOptions)
   If Len(sOptions) = 0 Then
      chkAuto.Value = vbChecked
   Else
      chkAuto.Value = Val(sOptions)
   End If
   
   'determine whether to enter elapsed time
   sOptions = GetSetting("Esi2000", "EsiAdmn", "ElapsedTime", sOptions)
   If Len(sOptions) = 0 Then
      chkElapsed.Value = vbUnchecked
   Else
      chkElapsed.Value = Val(sOptions)
   End If
   TurnElapsedTimeOnOrOff
   
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiAdmn", "AutoTime", chkAuto.Value
   SaveSetting "Esi2000", "EsiAdmn", "ElapsedTime", chkElapsed.Value
   
End Sub

'Private Sub cboFind_Change()
'    If cboFind.ListIndex >= 0 Then
'        If optTyp(iCurrIndex) = 1 Then
'                txtMon(iCurrIndex) = Trim(Left(cboFind.List(cboFind.ListIndex), 30))
'                txtRun(iCurrIndex) = Val(Right(cboFind.List(cboFind.ListIndex), 5))
'            End If
'        Else
'            Dim n As Integer
'            n = InStr(cboFind.List(cboFind.ListIndex), Chr(9))
'            If n > 0 Then
'                txtMon(iCurrIndex).Text = Left(cboFind.List(cboFind.ListIndex), n - 1)
'            End If
'        End If
'    End If
'End Sub

Private Sub cboFind_Click()
   If cboFind.ListIndex >= 0 Then
      If optTyp(iCurrIndex) = 1 Then
         txtMon(iCurrIndex) = Trim(Left(cboFind.List(cboFind.ListIndex), 30))
         
         'get run number from right hand side of entry
         txtRun(iCurrIndex) = Val(Right(cboFind.List(cboFind.ListIndex), Len(cboFind.List(cboFind.ListIndex)) - 30))
         vTimeCard(iCurrIndex + iIndex, TC_MO) = txtMon(iCurrIndex)
         vTimeCard(iCurrIndex + iIndex, TC_RUN) = txtRun(iCurrIndex)
      
      Else
         Dim n As Integer
         n = InStr(cboFind.List(cboFind.ListIndex), " ")
         If n > 0 Then
            txtMon(iCurrIndex).Text = Left(cboFind.List(cboFind.ListIndex), n - 1)
         End If
      End If
   End If
   
   txtMon(iCurrIndex).SetFocus
End Sub

'Private Sub cboFind_KeyPress(KeyAscii As Integer)
'   Debug.Print "avoid ESI_KeyBd stuff"
'End Sub
'
Private Sub chkDelete_LostFocus(Index As Integer)
   vTimeCard(iIndex + Index, TC_DELETE) = chkDelete(Index).Value
End Sub

Private Sub chkElapsed_Click()
   TurnElapsedTimeOnOrOff
End Sub

Private Sub cmbCde_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub cmbCde_KeyPress(Index As Integer, KeyAscii As Integer)
   
   'try to find a selection in the combobox starting with the key typed
   Dim I As Integer
   For I = 0 To cmbCde(Index).ListCount - 1
      'Debug.Print cmbCde(Index).List(i)
      If UCase(Left(cmbCde(Index).List(I), 1)) = UCase(Chr(KeyAscii)) Then
         cmbCde(Index).ListIndex = I
         Exit For
      End If
   Next
   KeyLock KeyAscii
   
End Sub


Private Sub cmbCde_LostFocus(Index As Integer)
   If Len(Trim(cmbCde(Index))) = 0 Then
      'Beep
      If cmbCde(Index).ListCount > 0 Then cmbCde(Index) = cmbCde(Index).List(0)
   End If
   vTimeCard(Index + iIndex, TC_TYPE) = cmbCde(Index)
   
End Sub


Private Sub cmbEmp_Click()
   bGoodEmployee = GetEmployee()
End Sub


Private Sub cmbEmp_LostFocus()
   cmbEmp = CheckLen(cmbEmp, 6)
   If Len(cmbEmp) Then
      cmbEmp = Format(cmbEmp, "000000")
      bGoodEmployee = GetEmployee()
   End If
   'optOhr.Enabled = True
   optInd.Enabled = True
   optInd.Caption = "&M__"
   bAddingCard = 0
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   bCanceled = 1
   cmbEmp = ""
   
End Sub


Private Sub cmdDn_Click()
   iCurrPage = iCurrPage + 1
   GetNextGroup
   
End Sub

Private Sub cmdDn_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub cmdEnd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If Val(lblTot) > 0 Then
      sMsg = "Do You Really Want To Cancel The Entry Of " & vbCrLf _
             & "This Time Card For " & lblNme & "?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         bAddingCard = 0
         cmdOps.Visible = False
         cmbEmp.Enabled = True
         txtDte.Enabled = True
         cmdCan.Enabled = True
         ResetBoxes
         On Error Resume Next
         cmbEmp.SetFocus
      Else
         CancelTrans
      End If
   Else
      bAddingCard = 0
      cmdOps.Visible = False
      cmdCan.Enabled = True
      cmbEmp.Enabled = True
      txtDte.Enabled = True
      ResetBoxes
      On Error Resume Next
      cmbEmp.SetFocus
   End If
   
End Sub

Private Sub cmdFind_Click()
   '    txtRns.Visible = True
   '    cmdOps.Visible = False
   '    On Error Resume Next
   '    txtRns.Text = txtMon(iCurrIndex).Text
   '    txtRns.SetFocus
   
   'display direct options
   Dim sFind As String
   sFind = txtMon(iCurrIndex).Text
   If Len(Trim(sFind)) < 1 Then
      MsgBox "Enter At Least (1) Leading Character.", _
         vbInformation, Caption
      txtMon(iCurrIndex).SetFocus
      Exit Sub
   End If
   
   If optTyp(iCurrIndex).Value = 1 Then
      FillTimeRuns sFind
   Else
      FillAccounts sFind
   End If
   
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor CURSOR_HOURGLASS
      OpenHelpContext "hs1502"
      cmdHlp = False
      MouseCursor CURSOR_NORMAL
   End If
   
End Sub



Private Sub cmdOps_Click()
   Dim b As Byte
   If cmdOps Then
      If Val(txtOpn(iCurrIndex)) > 0 Then
         If Trim(txtMon(iCurrIndex)) <> "" Or Val(txtRun(iCurrIndex)) <> 0 Then
            b = GetMoOperation(Compress(txtMon(iCurrIndex)), Val(txtRun(iCurrIndex)), _
                Val(txtOpn(iCurrIndex)))
            If b = 1 Then
               diaHmops.lblPrt = txtMon(iCurrIndex)
               diaHmops.lblRun = txtRun(iCurrIndex)
               diaHmops.txtOpn = Format(txtOpn(iCurrIndex), "000")
               diaHmops.Show
            Else
               MsgBox "MO Operation Wasn't Found, Is Canceled Or Closed.", _
                  vbInformation, Caption
            End If
         End If
      Else
         MsgBox "Please Enter The Operation Number.", _
            vbInformation, Caption
      End If
      cmdOps = False
   End If
   
End Sub

Private Sub cmdUpdate_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
    ' Make sure this is a valid employee
   If (Not bGoodEmployee) Then
    On Error Resume Next
    MsgBox "Not a Current Employee.", vbInformation, Caption
    cmbEmp.Enabled = True
    txtDte.Enabled = True
    cmbEmp.SetFocus
    Exit Sub
   End If
   
   
   'make sure time journal open for this date
   Dim bFound As Boolean
   Dim tc As New ClassTimeCharge
   bFound = tc.GetOpenTimeJournalForThisDate(txtDte.Text, sJournalID)
   If Not bFound Then
      Exit Sub
   End If
  
   'Diagnose "cmdUpdate_Click"
       
   If Val(lblTot) > 0 Then
      
      'first make sure all operations are valid
      sMsg = "Are You Ready To Update This " & vbCrLf _
             & "Time Card For " & lblNme & "?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         UpdateTimeCard
      Else
         CancelTrans
      End If
   Else
      MsgBox "There Are No Hours To Enter.", vbInformation, Caption
   End If
   On Error Resume Next
   cmbEmp.Enabled = True
   txtDte.Enabled = True
   cmbEmp.SetFocus
   
End Sub

Private Sub cmdUp_Click()
   iCurrPage = iCurrPage - 1
   GetNextGroup
   
End Sub

Private Sub cmdUp_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub cmdVew_Click()
   Dim iList As Integer
   Dim iRows As Integer
   MouseCursor CURSOR_HOURGLASS
   iRows = 10
   With diaTmvew.Grd
      .Rows = iRows
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      If Screen.Width > 9999 Then
         .ColWidth(0) = 400 * 1.25
         .ColWidth(1) = 1550 * 1.25
         .ColWidth(2) = 900 * 1.25
         .ColWidth(3) = 800 * 1.25
         .ColWidth(4) = 800 * 1.25
      Else
         .ColWidth(0) = 400
         .ColWidth(1) = 1550
         .ColWidth(2) = 900
         .ColWidth(3) = 800
         .ColWidth(4) = 800
      End If
   End With
'   For iList = 0 To LINES_PER_PAGE * MAX_PAGES - 1
    For iList = 0 To UBound(vTimeCard)
      If Len(Trim(vTimeCard(iList, TC_HOURS))) > 0 Then
         iRows = iRows + 1
         diaTmvew.Grd.Rows = iRows
         diaTmvew.Grd.row = iRows - 11
         
         diaTmvew.Grd.Col = 0
         If vTimeCard(iList, TC_DI) = 1 Then
            diaTmvew.Grd = "D"
         Else
            diaTmvew.Grd = "iList"
         End If
         diaTmvew.Grd.Col = 1
         diaTmvew.Grd = "" & vTimeCard(iList, TC_MO)
         
         diaTmvew.Grd.Col = 2
         diaTmvew.Grd = vTimeCard(iList, TC_BEGIN)
         
         diaTmvew.Grd.Col = 3
         diaTmvew.Grd = vTimeCard(iList, TC_END)
         
         diaTmvew.Grd.Col = 4
         diaTmvew.Grd = Format(Val(vTimeCard(iList, TC_HOURS)), "##0.000")
      End If
   Next
   diaTmvew.Show
   MouseCursor CURSOR_NORMAL
   
End Sub


Private Sub Form_Activate()
   Dim b As Byte
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
        ReDim vTimeCard(LINES_PER_PAGE * iTotalPages, 19)
        ReDim bGood(LINES_PER_PAGE * iTotalPages)

      Dim tc As New ClassTimeCharge
      sCoTimeAcct = tc.GetDefTimeAccounts("Time")
      sCoLaborAcct = tc.GetDefTimeAccounts("Labor")
      FillEmployees
      
      '        Dim bFound As Boolean
      '        bFound = GetOpenTimeJournalForThisDate(txtDte.Text)
      
      '        sJournalID = GetOpenJournal("TJ", Format$(ES_SYSDATE, "mm/dd/yy"))
      '        If Left(sJournalID, 4) = "None" Then
      '            sJournalID = ""
      '            b = 1
      '        Else
      '            If sJournalID = "" Then b = 0 Else b = 1
      '        End If
      '        If b = 0 Then
      '            MsgBox "There Is No Open Time Journal For This Period.", _
      '                vbExclamation, Caption
      '            Sleep 500
      '            Unload Me
      '            Exit Sub
      '        End If
      '
      bOnLoad = 0
   End If
   MouseCursor CURSOR_NORMAL
   
End Sub

Private Sub Form_Load()
   iTotalPages = 8
   FormLoad Me
   FormatControls
   
   'PREMTERMDT IS NULL AND
   sSql = "SELECT * FROM EmplTable WHERE PREMNUMBER = ? "
   'Set RdoQry1 = RdoCon.CreateQuery("", sSql)
   'RdoQry1.MaxRows = 1
   
    Set cmdObj1 = New ADODB.Command
    cmdObj1.CommandText = sSql
    
    Set prmObj1 = New ADODB.Parameter
    prmObj1.Type = adInteger
    cmdObj1.Parameters.Append prmObj1
   
   
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,PARTREF,PARTNUM " _
          & "FROM RunsTable,PartTable WHERE RUNREF=PARTREF AND (RUNREF= ? AND RUNNO = ? )"
   
   'Set RdoQry2 = RdoCon.CreateQuery("", sSql)
   'RdoQry2.MaxRows = 1

    Set cmdObj2 = New ADODB.Command
    cmdObj2.CommandText = sSql
    
   Set prmObj2 = New ADODB.Parameter
   prmObj2.Type = adChar
   prmObj2.Size = 30
   cmdObj2.Parameters.Append prmObj2
    
    Set prmObj3 = New ADODB.Parameter
    prmObj3.Type = adInteger
    cmdObj2.Parameters.Append prmObj3
   
   txtDte = Format(Now - 1, "mm/dd/yy")
   If sCurrDate = "" Then
      If Format(txtDte, "w") = 1 Then
         txtDte = Format(Now - 2, "mm/dd/yy")
      End If
   Else
      txtDte = sCurrDate
   End If
   'If sServer = "ESI_DEV_SVR" Then txtDte = "08/15/01"
   GetWeekEnd
   GetOptions
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   sCurrDate = txtDte
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set cmdObj1 = Nothing
   Set cmdObj2 = Nothing

   Set prmObj1 = Nothing
   Set prmObj2 = Nothing
   Set prmObj3 = Nothing

   Set diaHrtme = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   chkAuto.ForeColor = Es_FormBackColor
   chkElapsed.ForeColor = Es_FormBackColor
   optInd.ForeColor = Es_FormBackColor
   If sCurrDate = "" Then
      If Format(txtDte, "w") = 1 Then
         txtDte = Format(Now - 2, "mm/dd/yy")
      End If
   Else
      txtDte = sCurrDate
   End If
   cmdEnd.ToolTipText = "Cancel Work Not Updated And Return To Selection"
   cmdUp.ToolTipText = "Last Page (Page Up)"
   cmdDn.ToolTipText = "Next Page (Page Down)"
   
End Sub

Private Function GetCardNumber() As Byte
   'read timecard from db if it exists, otherwise, create a new card #
   'RETURN = True if successful
   
   Dim RdoCrd As ADODB.Recordset
   Dim iList As Integer
   
   Dim iMaxNumber As Integer
   
   On Error GoTo DiaErr1
   sOldCard = ""
   Erase vTimeCard
   iTotalPages = 8
   iMaxNumber = LINES_PER_PAGE * iTotalPages - 1
   'ReDim vTimeCard(0 To iMaxNumber, 16)
   
   
   'set time card to defaults
'   For iList = 0 To LINES_PER_PAGE * MAX_PAGES - 1
   ' For iList = 0 To UBound(vTimeCard, 1)
   '   vTimeCard(iList, TC_DI) = 1
   '   vTimeCard(iList, TC_MO) = ""
   '   vTimeCard(iList, TC_ACCT) = ""
   '   vTimeCard(iList, TC_RUN) = ""
   '   vTimeCard(iList, TC_OP) = ""
   '   vTimeCard(iList, TC_TYPE) = ""
   '   vTimeCard(iList, TC_BEGIN) = ""
   '   vTimeCard(iList, TC_END) = ""
   '   vTimeCard(iList, TC_HOURS) = ""
   '   vTimeCard(iList, TC_SRI) = ""
   '   vTimeCard(iList, TC_WC) = ""
   '   vTimeCard(iList, TC_TIMEACCT) = ""
   '   vTimeCard(iList, TC_DELETE) = 0
   '   vTimeCard(iList, TC_ACCEPTED) = 0
   '   vTimeCard(iList, TC_REJECTED) = 0
   '   vTimeCard(iList, TC_SCRAPPED) = 0
   'Next
   
   iList = -1
   
   'read card from db if it already exists
   MouseCursor CURSOR_HOURGLASS
   sSql = "SELECT TMCARD,TMEMP,TMDAY FROM TchdTable WHERE " _
          & "TMEMP=" & Val(cmbEmp) & " AND TMDAY='" & txtDte & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCrd, ES_FORWARD)
   If bSqlRows Then
      With RdoCrd
         sOldCard = "" & Trim(!TMCARD)
         .Cancel
      End With
'      sSql = "SELECT * FROM TcitTable WHERE TCCARD='" & sOldCard & "'" & vbCrLf _
'         & "order by case when len(TCSTART) < 6 then '' else substring(TCSTART,6,1)end, TCSTART"
      ' Order by tcStartTime
      sSql = "SELECT * FROM TcitTable WHERE TCCARD='" & sOldCard & "'" & vbCrLf _
         & " ORDER BY tcstarttime"
         
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCrd, ES_STATIC)
      
      If RdoCrd.RecordCount > iMaxNumber Then iMaxNumber = ((RdoCrd.RecordCount \ 8) * 8) + 8
    
      
      ReDim vTimeCard(0 To iMaxNumber, 18)
      ReDim bGood(0 To iMaxNumber)
      ClearArray
      iTotalPages = (RdoCrd.RecordCount \ LINES_PER_PAGE) + 1
      
      If bSqlRows Then
         With RdoCrd
            Do Until .EOF
               iList = iList + 1
               If !TCRUNNO > 0 Then
                  vTimeCard(iList, TC_DI) = 1
               Else
                  vTimeCard(iList, TC_DI) = 0
               End If
               vTimeCard(iList, TC_MO) = "" & Trim(!TCPARTREF)
               vTimeCard(iList, TC_ACCT) = "" & Trim(!TCACCT)
               vTimeCard(iList, TC_RUN) = "" & Trim(!TCRUNNO)
               vTimeCard(iList, TC_OP) = "" & Trim(!TCOPNO)
               vTimeCard(iList, TC_TYPE) = "" & Trim(!TCCODE)
               vTimeCard(iList, TC_BEGIN) = "" & Trim(!TCSTART)
               vTimeCard(iList, TC_END) = "" & Trim(!TCSTOP)
               vTimeCard(iList, TC_HOURS) = "" & Trim(!TCHOURS)
               vTimeCard(iList, TC_SRI) = "" & Trim(!TCSURUN)
               vTimeCard(iList, TC_WC) = "" & Trim(!TCWC)
               vTimeCard(iList, TC_TIMEACCT) = "" & Trim(!TCACCOUNT)
               vTimeCard(iList, TC_DELETE) = 0
               vTimeCard(iList, TC_ACCEPTED) = !TCACCEPT                'from POM
               vTimeCard(iList, TC_REJECTED) = !TCREJECT
               vTimeCard(iList, TC_SCRAPPED) = !TCSCRAP
               vTimeCard(iList, TC_STARTTIME) = !TCSTARTTIME
               vTimeCard(iList, TC_ENDTIME) = !TCSTOPTIME
               vTimeCard(iList, TC_COMMENTS) = Trim(!TCCOMMENTS)
               bGood(iList) = True
               .MoveNext
            Loop
            bAddingCard = 1
         End With
      End If
   Else
      ReDim vTimeCard(0 To iMaxNumber, 18)
      ReDim bGood(0 To iMaxNumber)
      iTotalPages = 1
      ClearArray
    
      'new time card
      'create a unique time card number
      'Stored as an (11) char string
      'The First (5) is the card date
      'the last (6) pickup the time to the part of a
      'second.
      'Find the Weekending based on WEEKENDS in ComnTable
      Dim S As Single
      Dim l As Long
      Dim m As Long
      Dim t As String
      
      '        m = DateValue(Format(ES_SYSDATE, "yyyy,mm,dd"))
      '        s = TimeValue(Format(ES_SYSDATE, "hh:mm:ss"))
      '        l = s * 1000000
      '        sCardNumber = Format(m, "00000") & Format(l, "000000")
      sCardNumber = GetNewNumber()
      
      iIndex = 0
      cmdDn.Enabled = True
      cmdDn.Picture = Endn
      GetCardNumber = True
      iCurrPage = 1
      lblPge = Format(iCurrPage, "#0")
      GetWeekEnd
      bAddingCard = 1
      MouseCursor CURSOR_NORMAL
      Exit Function
   End If
   If iList >= 0 Then
      iCurrPage = 1
      GetCardTime
      GetNextGroup
   End If
   '    Else
   '        MouseCursor CURSOR_NORMAL
   '            bAddingCard = 0
   '            If bCanceled = 0 Then
   '                sSql = "There Is No Time Card Recorded " & vbCrLf _
   '                    & "For " & Trim(lblNme) & " On " & txtDte & "."
   '                'Beep
   '                MsgBox sSql, vbInformation, Caption
   '            End If
   '        GetCardNumber = False
   '        On Error Resume Next
   '        cmbEmp.SetFocus
   '        Exit Function
   '    End If
   
   iIndex = 0
   iCurrPage = 1
   'cmdDn.Enabled = True
   'cmdDn.Picture = Endn
   GetCardNumber = True
   lblPge = Format(iCurrPage, "#0")
   GetWeekEnd
   MouseCursor CURSOR_NORMAL
   Exit Function
   
DiaErr1:
   sProcName = "getcardnu"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   MouseCursor CURSOR_NORMAL
End Function

'Private Sub lstRns_DblClick()
'    On Error Resume Next
'    If optTyp(iCurrIndex) = 1 Then
'        If lstRns.ListIndex >= 0 Then
'            txtMon(iCurrIndex) = Trim(Left(lstRns.List(lstRns.ListIndex), 30))
'            txtRun(iCurrIndex) = Val(Right(lstRns.List(lstRns.ListIndex), 5))
'        End If
'    Else
'        Dim n As Integer
'        n = InStr(lstRns.List(lstRns.ListIndex), Chr(9))
'        If n > 0 Then
'            txtMon(iCurrIndex).Text = Left(lstRns.List(lstRns.ListIndex), n - 1)
'        End If
'    End If
'    txtMon(iCurrIndex).SetFocus
'    lstRns.Height = 285
'    cmdOps.Visible = True
'    lstRns.Visible = False
'    txtRns.Visible = False
'
'End Sub


'Private Sub lstRns_KeyPress(KeyAscii As Integer)
'    On Error Resume Next
'    If KeyAscii = 13 Then
'        If lstRns.ListIndex >= 0 Then
'            txtMon(iCurrIndex) = Left(lstRns.List(lstRns.ListIndex), 29)
'            txtRun(iCurrIndex) = Val(Right(lstRns.List(lstRns.ListIndex), 5))
'        End If
'        lstRns.Height = 285
'        lstRns.Visible = False
'        txtRns.Visible = False
'        cmdOps.Visible = True
'    End If
'
'End Sub
'

Private Sub optInd_Click()
   Dim iList As Integer
   For iList = 0 To LINES_PER_PAGE - 1
      If optInd.Value = vbUnchecked Then
         optTyp(iList).Value = vbChecked
      Else
         optTyp(iList).Value = vbUnchecked
      End If
   Next
   
End Sub

Private Sub optInd_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optNew_Click()
   'never visible from New
   
End Sub

Private Sub optTyp_Click(Index As Integer)
   cboFind.Clear
   If optTyp(Index).Value = vbChecked Then
      'Direct
      txtMon(Index).Width = 2655
      txtRun(Index).Visible = True
      txtOpn(Index).Visible = True
      txtMon(Index) = " "
      txtSri(Index) = "R"
      vTimeCard(Index + iIndex, TC_DI) = 1
      vTimeCard(Index + iIndex, TC_SRI) = "R"
      vTimeCard(Index + iIndex, TC_ACCT) = ""
   Else
      'Indirect
      txtMon(Index).Width = 1340
      txtRun(Index).Visible = False
      txtOpn(Index).Visible = False
      txtSri(Index) = "I"
      txtMon(Index) = " "
      txtRun(Index) = " "
      txtOpn(Index) = " "
      vTimeCard(Index + iIndex, TC_DI) = 0
      vTimeCard(Index + iIndex, TC_SRI) = "I"
      vTimeCard(Index + iIndex, TC_MO) = ""
      vTimeCard(Index + iIndex, TC_RUN) = ""
      vTimeCard(Index + iIndex, TC_OP) = ""
      vTimeCard(Index + iIndex, TC_WC) = ""
   End If
   
End Sub

Private Sub optTyp_GotFocus(Index As Integer)
   If bAddingCard = 1 Then
      optInd.Enabled = False
      'optOhr.Enabled = False
      optInd.Caption = "___"
      cmbEmp.Enabled = False
      txtDte.Enabled = False
      cmdCan.Enabled = False
   Else
      bAddingCard = 0
   End If
   
End Sub


Private Sub optTyp_KeyPress(Index As Integer, KeyAscii As Integer)
   'KeyLock KeyAscii
   
End Sub


Private Sub optTyp_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtBeg_GotFocus(Index As Integer)
   SelectFormat Me
   ChangeRow Index
End Sub


Private Sub txtBeg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtBeg_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub


Private Sub txtBeg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

'Private Sub txtBeg_LostFocus(Index As Integer)
'   Dim dBegTime As Date
'   Dim dEndTime As Date
'
'   'On Error Resume Next
'   If Len(Trim(txtBeg(Index))) Then
'      txtBeg(Index) = GetTime(txtBeg(Index))
'
'      'if time not changed, don't continue
'      If DateDiff("n", txtBeg(Index), vTimeCard(Index + iIndex, TC_BEGIN)) = 0 Then
'         Exit Sub
'      End If
'
'      If Index > 0 Then
'         dEndTime = txtbeg(Index - 1)
'         dBegTime = txtBeg(Index)
'         If TimeValue(dBegTime) < TimeValue(dEndTime) Then
'            If chkAuto.Value = vbChecked Then
'               dBegTime = Format(dBegTime + 0.5, "hh:nna/p")
'               txtBeg(Index) = Format(dBegTime, "hh:nna/p")
'            End If
'         End If
'      End If
'   End If
'   vTimeCard(Index + iIndex, TC_BEGIN) = txtBeg(Index)
'
'End Sub
'

Private Sub txtBeg_LostFocus(Index As Integer)
   
   'if done editing, just return
   If Not cmdUpdate.Enabled Then
      Exit Sub
   End If

   'if invalid time, don't continue
   Dim tc As New ClassTimeCharge
   txtBeg(Index) = tc.GetTime(txtBeg(Index))    'returns blank if invalid
   If txtBeg(Index) = "" Then
      ThereIsAnError Index
      Me.txtHrs(Index) = ""
      Exit Sub
   End If
   ThereIsNoError Index
   
   bGoodTime = CheckTime(Index, True)
   If bGoodTime Then
      vTimeCard(Index + iIndex, TC_BEGIN) = txtBeg(Index)
   End If
End Sub


Private Sub txtDte_DropDown()
'   If MDISect.ActiveForm.ActiveControl.Name <> "txtDte" Then
'      MsgBox MDISect.ActiveForm.ActiveControl.Name
'   End If
   ShowCalendar Me
End Sub

'Private Function GetOpenTimeJournalForThisDate(dt As Variant) As Boolean
'   'place it in sJournalID
'   'RETURN = True if successful
'
'   Dim b As Boolean
'
'   sJournalID = GetOpenJournal("TJ", Format$(dt, "mm/dd/yy"))
'   If Left(sJournalID, 4) = "None" Then
'      sJournalID = ""
'      b = True
'   Else
'      If sJournalID = "" Then b = False Else b = True
'   End If
'   If Not b Then
'      MsgBox "There Is No Open Time Journal For This Period.", _
'         vbExclamation, Caption
'      Sleep 500
'      'Unload Me
'   End If
'   GetOpenTimeJournalForThisDate = b
'End Function
'
'Get the employees and other stuff that we'll need

Private Sub FillEmployees()
   Dim RdoCmb As ADODB.Recordset
   Dim iList As Integer
   
   On Error GoTo DiaErr1
   'sSql = "Qry_FillEmployees"
    sSql = "select PREMNUMBER from EmplTable where (  (PREMTERMDT IS NULL) or (PREMTERMDT IS NOT NULL AND PREMREHIREDT > PREMTERMDT) ) AND ( PREMSTATUS NOT IN ('D','I'))" _
          & " order by PREMNUMBER"
   LoadNumComboBox cmbEmp, "000000"
   If bSqlRows Then cmbEmp = cmbEmp.List(0)
   
   On Error Resume Next
   iTotalCodes = -1
   sSql = "SELECT TYPECODE,TYPEADDER FROM TmcdTable ORDER BY TYPESEQ "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            iTotalCodes = iTotalCodes + 1
            vTypeCode(iTotalCodes, 0) = "" & Trim(!typeCode)
            vTypeCode(iTotalCodes, 1) = !TYPEADDER
            For iList = 0 To LINES_PER_PAGE - 1
               AddComboStr cmbCde(iList).hwnd, "" & Trim(!typeCode)
            Next
            .MoveNext
         Loop
         .Cancel
      End With
   Else
      sSql = "There Are No Time Codes. Please Select" & vbCrLf _
             & "Time Type Codes To Establish Defaults."
      MsgBox sSql, vbInformation, Caption
   End If
   If cmbCde(0).ListCount > 0 Then
      For iList = 0 To LINES_PER_PAGE - 1
         cmbCde(iList) = cmbCde(iList).List(0)
      Next
   End If
   
   'get workcenter overhead rates.  use rates from shops where not present
   sSql = "SELECT WCNREF, WCNOHPCT, WCNOHFIXED," & vbCrLf _
          & "SHPOHRATE, SHPOHTOTAL" & vbCrLf _
          & "FROM WcntTable" & vbCrLf _
          & "JOIN ShopTable on WCNSHOP = SHPREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      iTotalCenters = -1
      With RdoCmb
         Do Until .EOF
            iTotalCenters = iTotalCenters + 1
            vWorkCenter(iTotalCenters, WC_CODE) = "" & Trim(!WCNREF)
            vWorkCenter(iTotalCenters, WC_OHPERCENT) = !WCNOHPCT
            vWorkCenter(iTotalCenters, WC_OHFIXED) = !WCNOHFIXED
            
            If vWorkCenter(iTotalCenters, WC_OHPERCENT) = 0 _
                           And vWorkCenter(iTotalCenters, WC_OHFIXED) = 0 Then
               vWorkCenter(iTotalCenters, WC_OHPERCENT) = !SHPOHRATE
               vWorkCenter(iTotalCenters, WC_OHFIXED) = !SHPOHTOTAL
            End If
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoCmb = Nothing
   ResetBoxes
   If cmbEmp.ListCount > 0 Then
      If optNew.Value = vbUnchecked Then
         cmbEmp = cmbEmp.List(0)
      End If
      bGoodEmployee = GetEmployee()
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillemplo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetEmployee() As Byte
   On Error GoTo DiaErr1
   Dim RdoEmp As ADODB.Recordset
   Dim strTerm As String
   MouseCursor CURSOR_HOURGLASS
   cmdUp.Enabled = False
   cmdUp.Picture = Dsup
   'RdoQry1(0) = Val(cmbEmp)
   cmdObj1.Parameters(0).Value = Val(cmbEmp)
   bSqlRows = clsADOCon.GetQuerySet(RdoEmp, cmdObj1)
   'MsgBox "emp" & cmbEmp
   If bSqlRows Then
      With RdoEmp
         If (Not IsNull(!PREMTERMDT)) And (!PREMREHIREDT < !PREMTERMDT) Then
            MsgBox "Not a Current Employee.", vbInformation, Caption
            Set RdoEmp = Nothing
            GetEmployee = False
            lblNme = "Not a Current Employee"
            lblSsn = ""
            MouseCursor CURSOR_NORMAL
            Exit Function
         End If
         cmbEmp = Format(!PREMNUMBER, "000000")
         lblNme = "" & Trim(!PREMLSTNAME) & ", " _
                  & Trim(!PREMFSTNAME) & " " _
                  & Trim(!PREMMINIT)
         lblSsn = "" & Trim(!PREMSOCSEC)
         On Error Resume Next
         
         'if salaried, divide rate by # of hours per month
         If Trim(!PREMHOURLY) = "S" Then
            cEmplRate = Format(!PREMPAYRATE / 173.34, "#####0.00")
         Else
            'cEmplRate = Format(!PREMPAYRATE, ES_QuantityDataFormat)
            cEmplRate = Format(!PREMPAYRATE, "#####0.00")
         End If
         
         'MsgBox "cemplrate=" & cEmplRate
         sEmplCenter = "" & Trim(!PREMCENTER)
         sEmplCenter = Compress(sEmplCenter)
         
         sEmplShop = "" & Trim(!PREMSHOP)
         sEmplShop = Compress(sEmplShop)
         sEmplAcct = "" & Trim(!PREMACCTS)
         sEmplAcct = Compress(sEmplAcct)
         
         .Cancel
         GetEmployee = True
         sCurrEmployee = cmbEmp
      End With
   Else
      GetEmployee = False
      lblNme = "No Current Employee"
      lblSsn = ""
   End If
   Set RdoEmp = Nothing
   '    If optOhr.Value = vbChecked Then GetShopInfo _
   '        Else GetWorkCenterInfo
   
   ' if wc overhead = 0 then use shop overhead
   GetWorkCenterInfo
   If cShopOhRate = 0 And cShopOhFixed = 0 Then
      GetShopInfo
   End If
   MouseCursor CURSOR_NORMAL
   Exit Function
   
DiaErr1:
   sProcName = "getemploy"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   MouseCursor CURSOR_NORMAL
   
End Function

Private Sub txtDte_KeyDown(KeyCode As Integer, Shift As Integer)
'   Debug.Print KeyCode
   On Error GoTo whoops
   If KeyCode = vbKeyUp Then 'up
      txtDte.Text = DateAdd("d", 1, txtDte.Text)
   ElseIf KeyCode = vbKeyDown Then 'down
      txtDte.Text = DateAdd("d", -1, txtDte.Text)
   End If
   Exit Sub
whoops:
End Sub

Private Sub txtDte_LostFocus()
   Dim iList As Integer
   
   'don't check when just leaving field to get the date
   If Screen.ActiveForm.Name = "SysCalendar" Then
      Exit Sub
   End If
   
   'make sure there is a journal open for this date
   Dim bFound As Boolean
   On Error Resume Next
   Dim v As Variant
   v = DateValue(txtDte.Text)
   If Err Then
      Exit Sub
   End If
   On Error GoTo 0
   Dim JournalID As String
   Dim tc As New ClassTimeCharge
   If Not tc.GetOpenTimeJournalForThisDate(txtDte.Text, sJournalID) Then
      Exit Sub
   End If
   
   txtDte = CheckDate(txtDte)
   If DateValue(Format(txtDte, "yyyy,mm,dd")) > DateValue(Format(ES_SYSDATE, "yyyy,mm,dd")) Then
      
      Dim bResponse As Byte
      Dim sMsg As String
      
      sMsg = "Do you Really want to entry post dated Time Card? (" & Format(txtDte, "mm/dd/yy") & ")"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         txtDte = Format(ES_SYSDATE, "mm/dd/yy")
      End If
      'Beep
      'txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   Else
      bGoodCard = GetCardNumber()
      If bGoodCard Then
         For iList = 0 To LINES_PER_PAGE - 1
            optTyp(iList).Enabled = True
            txtMon(iList).Enabled = True
            txtRun(iList).Enabled = True
            txtOpn(iList).Enabled = True
            txtSri(iList).Enabled = True
            cmbCde(iList).Enabled = True
            txtBeg(iList).Enabled = (chkElapsed.Value = 0)
            txtEnd(iList).Enabled = (chkElapsed.Value = 0)
            txtHrs(iList).Enabled = (chkElapsed.Value = 1)
            chkDelete(iList).Enabled = True
            bAddingCard = 1
         Next
         cmdUpdate.Enabled = True
         cmdEnd.Enabled = True
         TurnElapsedTimeOnOrOff
         On Error Resume Next
      End If
   End If
   GetWeekEnd
   
End Sub


Private Sub txtend_GotFocus(Index As Integer)
   SelectFormat Me
   ChangeRow Index
End Sub


Private Sub txtEnd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
End Sub

Private Sub txtEnd_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
End Sub


Private Sub txtEnd_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
End Sub

Private Sub txtEnd_LostFocus(Index As Integer)
   
   'if done editing, just return
   If Not cmdUpdate.Enabled Then
      Exit Sub
   End If

   'if invalid time, don't continue
   Dim tc As New ClassTimeCharge
   txtEnd(Index) = tc.GetTime(txtEnd(Index))    'returns blank if invalid
   If txtEnd(Index) = "" Then
      ThereIsAnError Index
      Me.txtHrs(Index) = ""
      Exit Sub
   End If
   ThereIsNoError Index

''   ' if not adding charges to card (like after an update), don't continue
''   If bAddingCard = 0 Then
''      Exit Sub
''   End If

   bGoodTime = CheckTime(Index, False)
   If bGoodTime Then
      vTimeCard(Index + iIndex, TC_END) = txtEnd(Index)
      'If Val(Left(txtEnd(Index), 2)) > 0 Then
         If Index < LINES_PER_PAGE - 1 Then
            If Left(txtBeg(Index + 1), 1) = " " Then
               If chkAuto.Value = vbChecked Then txtBeg(Index + 1) = txtEnd(Index)
            End If
         Else
            EnableDisableUpDown
         End If
      'End If
   End If
End Sub

Private Function EnableDisableUpDown()
   If iIndex > 1 Then
      cmdUp.Enabled = True
      cmdUp.Picture = Enup
   Else
      cmdUp.Enabled = False
      cmdUp.Picture = Dsup
   End If
   
   'If Val(txtHrs(LINES_PER_PAGE - 1)) > 0 And iIndex < MAX_PAGES Then
   If Val(txtHrs(LINES_PER_PAGE - 1)) > 0 And iCurrPage < iTotalPages Then
      cmdDn.Enabled = True
      cmdDn.Picture = Endn
   Else
      cmdDn.Enabled = False
      cmdDn.Picture = Dsdn
   End If
End Function

Private Function CheckTime(Index2 As Integer, OkIfEndingTimeNotYetEntered As Boolean) As Boolean
   'processes time charge line
   'RETURN = true if successful
   '       = false if error
   
   Dim iList As Integer
   Dim cTotal As Currency
   CheckTime = False
   
   Dim tm As New ClassTimeCharge
   If OkIfEndingTimeNotYetEntered And Not tm.IsValidTime(txtEnd(Index2)) Then
      CheckTime = tm.IsValidTime(txtBeg(Index2))
      Exit Function
   End If
      
   On Error Resume Next
   
   'test start time & stop time.  update elapsed time
   If chkElapsed.Value = 0 Then
      'txtHrs(Index2) = Format((TimeValue(txtEnd(Index2)) - TimeValue(txtBeg(Index2))) * 24, "##0.00")
      Dim minutes As Integer
      minutes = DateDiff("n", txtBeg(Index2), txtEnd(Index2))
      If minutes < 0 Then
         minutes = minutes + 24 * 60
      End If
      txtHrs(Index2) = Format(minutes / 60, "##0.00")
      If Err = 0 Then
         txtHrs(Index2).ForeColor = Es_TextForeColor
         ThereIsNoError Index2
      Else
         If Trim(txtBeg(Index2)) <> "" Or Trim(txtEnd(Index2)) <> "" Then
            If Val(txtHrs(Index2)) < 0 Then
               ThereIsAnError Index2
               Exit Function
            End If
         End If
      End If
      
      'test elapsed time.  Set start time and stop time to midnight
   Else
      If Val(txtHrs(Index2)) > 0 Then
         txtBeg(Index2) = "0:00a"
         txtEnd(Index2) = "0:00a"
         ThereIsNoError Index2
      Else
         ThereIsAnError Index2
         Exit Function
      End If
   End If
   
   If Val(txtHrs(Index2)) <= 0 Then
      ThereIsAnError Index2
      Exit Function
   Else
      ThereIsNoError Index2
   End If
   
   'time info is valid.
   vTimeCard(iIndex + Index2, TC_DI) = optTyp(Index2).Value
   vTimeCard(iIndex + Index2, TC_TYPE) = cmbCde(Index2)
   vTimeCard(iIndex + Index2, TC_BEGIN) = txtBeg(Index2)
   vTimeCard(iIndex + Index2, TC_END) = txtEnd(Index2)
   vTimeCard(iIndex + Index2, TC_HOURS) = txtHrs(Index2)
   vTimeCard(iIndex + Index2, TC_SRI) = txtSri(Index2)
   
   'Check for valid run or account
   If optTyp(Index2).Value = vbChecked Then
      If Not TestRunOp(Index2) Then
         Exit Function
      End If
      vTimeCard(iIndex + Index2, TC_MO) = txtMon(Index2)
      vTimeCard(iIndex + Index2, TC_ACCT) = " "
   Else
      If Not TestAccount(Index2) Then
         Exit Function
      End If
      vTimeCard(iIndex + Index2, TC_MO) = " "
      vTimeCard(iIndex + Index2, TC_ACCT) = txtMon(Index2)
      vTimeCard(iIndex + Index2, TC_TIMEACCT) = txtMon(Index2)
   End If
   
   'charge is valid.  save it in the array
   CheckTime = True
   vTimeCard(iIndex + Index2, TC_RUN) = txtRun(Index2)
   vTimeCard(iIndex + Index2, TC_OP) = txtOpn(Index2)
'   For iList = 0 To LINES_PER_PAGE * MAX_PAGES - 1
   For iList = 0 To UBound(vTimeCard)
      cTotal = cTotal + Val(vTimeCard(iList, TC_HOURS))
   Next
   lblTot = Format(cTotal, "###0.000")
   
   'if last line on page, enable next page
   cmdDn.Enabled = (iCurrPage < iTotalPages And (Index2 + 1) Mod LINES_PER_PAGE = 0)
   If cmdDn.Enabled Then
      cmdDn.Picture = Endn
   Else
      cmdDn.Picture = Dsdn
   End If
   
End Function


'Private Sub txtEnd_Validate(Index As Integer, Cancel As Boolean)
'
'   'if invalid time, don't continue
'   txtEnd(Index) = GetTime(txtEnd(Index))    'returns blank if invalid
'   If txtEnd(Index) = "" Then
'      ThereIsAnError Index
'      Me.txtHrs(Index) = ""
'      Exit Sub
'   End If
'   ThereIsNoError Index
'
'''   ' if not adding charges to card (like after an update), don't continue
'''   If bAddingCard = 0 Then
'''      Exit Sub
'''   End If
'
'   bGoodTime = CheckTime(Index, False)
'   If bGoodTime Then
'      vTimeCard(Index + iIndex, TC_END) = txtEnd(Index)
'      'If Val(Left(txtEnd(Index), 2)) > 0 Then
'         If Index < LINES_PER_PAGE - 1 Then
'            If Left(txtBeg(Index + 1), 1) = " " Then
'               If chkAuto.Value = vbChecked Then txtBeg(Index + 1) = txtEnd(Index)
'            End If
'         Else
'            EnableDisableUpDown
'         End If
'      'End If
'   End If
'End Sub

Private Sub txtHrs_LostFocus(Index As Integer)
   ' if not adding charges to card (like after an update), don't continue
   ChangeRow Index
   If bAddingCard = 0 Then
      GetCardTime
      Exit Sub
   End If
   
   bGoodTime = CheckTime(Index, False)
   If bGoodTime Then
      vTimeCard(Index + iIndex, TC_END) = txtEnd(Index)
      If Val(Left(txtEnd(Index), 2)) > 0 Then
         If Index < LINES_PER_PAGE - 1 Then
            If Left(txtBeg(Index + 1), 1) = " " Then
               If chkAuto.Value = vbChecked Then txtBeg(Index + 1) = txtEnd(Index)
            End If
         Else
            EnableDisableUpDown
         End If
      End If
   End If
   'Diagnose ("Index = " & Index)
End Sub

Private Sub txtMon_GotFocus(Index As Integer)
   SelectFormat Me
   
   On Error Resume Next
   Dim v As Variant
   v = TimeValue(txtBeg(Index).Text)
   If Err Then
      'use previous end time as start time
      If iCurrPage > 1 Or Index > 0 Then
         txtBeg(Index).Text = vTimeCard((iCurrPage - 1) * LINES_PER_PAGE + Index - 1, TC_END)
      Else
         'Debug.Print "breakpoint"
      End If
   End If
   
   ChangeRow Index
   cmdOps.Visible = False
   cmdFind.Enabled = True
   
End Sub

Private Sub txtMon_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtMon_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub txtMon_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub Diagnose(sComment As String)
   Debug.Print ""
   Debug.Print sComment
   Debug.Print "######################"
   Dim iList As Integer
'   For iList = 0 To LINES_PER_PAGE * MAX_PAGES - 1
    For iList = 0 To UBound(vTimeCard)
      If bGood(iList) Or Val(vTimeCard(iList, TC_HOURS)) <> 0 Then
         Debug.Print iList & ": GOOD=" & bGood(iList) _
            & " DI=" & (vTimeCard(iList, TC_DI)) _
            & " MO " & (vTimeCard(iList, TC_MO)) _
            & "  # " & (vTimeCard(iList, TC_RUN)) _
            & " OP " & (vTimeCard(iList, TC_OP)) _
            & " ACCT=" & (vTimeCard(iList, TC_ACCT)) _
            & " " & (vTimeCard(iList, TC_TYPE)) _
            & " " & (vTimeCard(iList, TC_BEGIN)) _
            & "-" & (vTimeCard(iList, TC_END)) _
            & "=" & (vTimeCard(iList, TC_HOURS)) _
            & " " & (vTimeCard(iList, TC_SRI)) _
            & " WC=" & (vTimeCard(iList, TC_WC)) _
            & " TMACCT=" & (vTimeCard(iList, TC_TIMEACCT))
         If Trim(vTimeCard(iList, TC_ACCT)) = "" Or Trim(vTimeCard(iList, TC_TIMEACCT)) = "" Then
            Debug.Print "blank TC_ACCT='" & Trim(vTimeCard(iList, TC_ACCT)) & "' TC_ACCOUNT='" & Trim(vTimeCard(iList, TC_TIMEACCT)) & "'"
         End If
      End If
   Next
End Sub

Private Sub txtMon_LostFocus(Index As Integer)
   
   If optTyp(Index).Value = vbChecked Then
      txtMon(Index) = CheckLen(txtMon(Index), 30)
      
      'test for invalid part/run/op
      bGood(Index + iIndex) = TestRunOp(Index)
      vTimeCard(Index + iIndex, TC_MO) = txtMon(Index)
      If Not bGood(Index + iIndex) Then
         Exit Sub
      End If
      
      'vTimeCard(Index + iIndex, TC_MO) = txtMon(Index)
      vTimeCard(Index + iIndex, TC_ACCT) = ""
   Else
      'don't validate if doing an account search
      If Screen.ActiveControl.Name = "cmdFind" Then
         Exit Sub
      End If
      
      txtMon(Index) = CheckLen(txtMon(Index), 12)
      bGood(Index + iIndex) = TestAccount(Index)
      If Not bGood(Index + iIndex) Then
         Exit Sub
      End If
      vTimeCard(Index + iIndex, TC_MO) = ""
      vTimeCard(Index + iIndex, TC_ACCT) = txtMon(Index)
      vTimeCard(Index + iIndex, TC_TIMEACCT) = txtMon(Index)
   End If
   ChangeRow Index
   
   'test for invalid part/run/op
   bGood(Index + iIndex) = TestRunOp(Index)
End Sub


Private Sub txtOpn_GotFocus(Index As Integer)
   SelectFormat Me
   If txtMon(Index) <> "" And Val(txtRun(Index)) > 0 Then cmdOps.Visible = True
   ChangeRow Index
End Sub

Private Sub txtOpn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtOpn_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtOpn_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtOpn_LostFocus(Index As Integer)
   'test for invalid part/run/op
   bGood(Index + iIndex) = TestRunOp(Index)
   If Not bGood(Index + iIndex) Then
      Exit Sub
   End If
   
   txtOpn(Index) = CheckLen(txtOpn(Index), 4)
   txtOpn(Index) = Format(Abs(Val(txtOpn(Index))), "000")
   vTimeCard(Index + iIndex, TC_OP) = txtOpn(Index)
   GetRunOp Index
   
End Sub

Private Sub txtRns_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = Asc(vbCr) Then
      cmdFind_Click
   End If
End Sub

Private Sub txtRun_GotFocus(Index As Integer)
   ChangeRow Index
End Sub

'Private Sub txtRns_LostFocus()
'
'    'display direct options
'    If optTyp(iCurrIndex).Value = 1 Then
'        FillTimeRuns
'    Else
'        FillAccounts
'    End If
'
'End Sub
'
'
'Private Sub txtRun_GotFocus(Index As Integer)
'    SelectFormat Me
'    cmdFind.Enabled = False
'    If lstRns.Visible Then
'        lstRns.Visible = False
'        txtRns.Visible = False
'    End If
'
'End Sub
'

Private Sub txtRun_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtRun_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtRun_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

'Private Sub txtRun_LostFocus(Index As Integer)
'    txtRun(Index) = CheckLen(txtRun(Index), 5)
'    txtRun(Index) = Format(Abs(Val(txtRun(Index))), "####0")
'    vTimeCard(Index + iIndex, TC_RUN) = txtRun(Index)
'    bGood(Index + iIndex) = GetRun(Index)
'    cmdFind.Enabled = False
'    If lstRns.Visible Then
'        txtRns.Height = 300
'        lstRns.Visible = False
'        txtRns.Visible = False
'    End If
'    If txtMon(Index) <> "" Or Val(txtRun(Index)) > 0 Then cmdOps.Visible = True
'
'End Sub

Private Sub GetNextGroup()
   
   Dim b As Byte
   Dim I As Integer
   Dim iList As Integer
   
   cmbEmp.Enabled = False
   txtDte.Enabled = False
   
   MouseCursor CURSOR_HOURGLASS
   On Error Resume Next
   
   If iCurrPage < 1 Then iCurrPage = 1
   If iCurrPage = 1 Then
      cmdUp.Picture = Dsup
      cmdDn.Picture = Endn
   Else
      cmdUp.Picture = Enup
   End If
   If iCurrPage > iTotalPages Then iCurrPage = iTotalPages
   If iCurrPage = iTotalPages Then
      cmdUp.Picture = Enup
      cmdDn.Picture = Dsdn
   Else
      cmdDn.Picture = Endn
   End If
   lblPge = Format(iCurrPage, "#0")
   iIndex = 8 * (iCurrPage - 1)
   
   I = 0
   For iList = iIndex To iIndex + LINES_PER_PAGE - 1
      optTyp(I).Value = Val(vTimeCard(iList, TC_DI))
      chkDelete(I).Value = Val(vTimeCard(iList, TC_DELETE))
      If Val(vTimeCard(iList, TC_DI)) = 1 Then
         txtMon(I) = vTimeCard(iList, TC_MO)
         txtRun(I) = vTimeCard(iList, TC_RUN)
         txtOpn(I) = vTimeCard(iList, TC_OP)
         b = GetRun(I)
      Else
         txtMon(I) = vTimeCard(iList, TC_ACCT)
      End If
      cmbCde(I) = vTimeCard(iList, TC_TYPE)
      If Trim(vTimeCard(iList, TC_SRI)) <> "" Then
         txtSri(I) = vTimeCard(iList, TC_SRI)
      Else
         If optTyp(I).Value = vbChecked Then
            txtSri(I) = "R"
         Else
            txtSri(I) = "iList"
         End If
      End If
      
      
      If Val(vTimeCard(iList, TC_HOURS)) >= 0 Then
         txtBeg(I) = vTimeCard(iList, TC_BEGIN)
         txtEnd(I) = vTimeCard(iList, TC_END)
         ThereIsNoError I
         txtHrs(I) = Format(vTimeCard(iList, TC_HOURS), "##0.000")
         
         'if begin time = end time but hours > 0, make sure elapsed time is selected
         If (txtBeg(I) = txtEnd(I)) And (Val(txtHrs(I)) > 0) And (chkElapsed.Value = 0) Then
            chkElapsed.Value = 1
         End If
      Else
         If vTimeCard(iList, TC_HOURS) = "* Error *" Then
            ThereIsAnError I
            txtHrs(I) = vTimeCard(iList, TC_HOURS)
            txtBeg(I) = vTimeCard(iList, TC_BEGIN)
            txtEnd(I) = vTimeCard(iList, TC_END)
         Else
            ThereIsNoError I
            txtBeg(I) = "  :  "
            txtEnd(I) = "  :  "
            txtHrs(I) = ""
         End If
         If Not bGood(iList) Then
            If Len(Trim(txtMon(I))) > 0 Then
               txtMon(I).ForeColor = ES_RED
            Else
               txtMon(I).ForeColor = Es_TextForeColor
            End If
         Else
            txtMon(I).ForeColor = Es_TextForeColor
         End If
      End If
      I = I + 1
      'If i > 7 Then Exit For     'redundant
   Next
   
   If iIndex > 0 Then
      If Val(vTimeCard(iIndex - 1, TC_HOURS)) > 0 Then
         If Val(Left(txtBeg(0), 2)) = 0 Then txtBeg(0) = vTimeCard(iIndex - 1, TC_END)
      End If
   End If
   
   'enable previous page if there is one
   '    If iIndex > 0 Then
   '        cmdUp.Enabled = True
   '    Else
   '        cmdUp.Enabled = False
   '    End If
   If iIndex > 0 Then
      cmdUp.Enabled = True
      cmdUp.Picture = Enup
   Else
      cmdUp.Enabled = False
      cmdUp.Picture = Dsup
   End If
   
   'enable next page if current page is full
   '    If iCurrPage < MAX_PAGES Then
   '        cmdDn.Enabled = Val(vTimeCard(iIndex + LINES_PER_PAGE - 1, TC_HOURS)) > 0
   '    Else
   '        cmdDn.Enabled = False
   '    End If
   '    If cmdDn.Enabled Then
   '        cmdDn.Picture = Endn
   '    Else
   '        cmdDn.Picture = Dsdn
   '    End If
   
   EnableDisableUpDown
   
   MouseCursor CURSOR_NORMAL
   
End Sub

Private Sub ResetBoxes()
   Dim iList As Integer
   lstItm.Clear
   For iList = 0 To LINES_PER_PAGE - 1
      If optInd.Value = vbUnchecked Then
         optTyp(iList).Value = vbChecked
      Else
         optTyp(iList).Value = vbUnchecked
      End If
      txtMon(iList) = " "
      txtMon(iList).ForeColor = Es_TextForeColor
      txtRun(iList) = " "
      txtOpn(iList) = " "
      txtSri(iList) = "R"
      If cmbCde(0).ListCount > 0 Then cmbCde(iList) = cmbCde(0).List(0)
      txtBeg(iList) = "  :  "
      txtEnd(iList) = "  :  "
      txtHrs(iList) = " "
      ThereIsNoError iList
      lblTot = "0.00"
      optTyp(iList).Enabled = False
      txtMon(iList).Enabled = False
      txtRun(iList).Enabled = False
      txtOpn(iList).Enabled = False
      txtSri(iList).Enabled = False
      cmbCde(iList).Enabled = False
      txtBeg(iList).Enabled = False
      txtEnd(iList).Enabled = False
      txtSri(iList).Enabled = False
      chkDelete(iList).Enabled = False
      'optTyp(iList).Value = vbChecked
   Next
   cmdCan.Enabled = True
   cmbEmp.Enabled = True
   txtDte.Enabled = True
   cmdUpdate.Enabled = False
   cmdEnd.Enabled = False
   Erase vTimeCard
'   Erase bGood
   ReDim vTimeCard(LINES_PER_PAGE * iTotalPages, 16)
'   ReDim bGood(LINES_PER_PAGE * MAX_PAGES)
   
   
'   For iList = 0 To LINES_PER_PAGE * MAX_PAGES - 1
   For iList = 0 To UBound(vTimeCard)
      vTimeCard(iList, TC_DI) = 1
      vTimeCard(iList, TC_HOURS) = 0
   Next
   
   cmdUp.Enabled = False
   cmdDn.Enabled = False
   
End Sub


Private Sub UpdateTimeCard()
   'Dim i As Integer
   Dim iList As Integer
   
   Dim iOpno As Integer
   Dim iRef As Integer
   Dim iRunno As Integer
   Dim iRows As Integer
   Dim typeCode As String
   Dim sBegTime As String
   Dim sEndTime As String
   Dim sPartNumber As String
   Dim sWorkCenter As String
   Dim sShop As String
   Dim sDebitAcct As String
   
   Dim bTimeOrg As Boolean
   On Error GoTo DiaErr1
   
   Dim tc As New ClassTimeCharge
   
   bAddingCard = 1
   
   'do some modest testing
'   For iList = 0 To LINES_PER_PAGE * MAX_PAGES - 1
'Debug.Print CStr(UBound(vTimeCard))
   For iList = 0 To UBound(vTimeCard)
      If Val(Trim(vTimeCard(iList, TC_HOURS))) > 0 Then
         
         'direct charge
         If Val(vTimeCard(iList, TC_DI)) = 1 Then
            'If Not bGood(iList)
            If Not IsValidRunOp(CStr(vTimeCard(iList, TC_MO)), _
                                CStr(vTimeCard(iList, TC_RUN)), CStr(vTimeCard(iList, TC_OP)), False, False) Then
               MsgBox "Invalid MO " & vTimeCard(iList, TC_MO) _
                  & " run " & vTimeCard(iList, TC_RUN) _
                  & " operation " & vTimeCard(iList, TC_OP), vbExclamation, Caption
               Exit Sub
            End If
            
         'indirect charge
         Else
            If Not IsValidAccount(Trim(vTimeCard(iList, TC_ACCT))) Then
               MsgBox "Indirect time charge has invalid account " & vTimeCard(iList, TC_ACCT), vbExclamation, Caption
               Exit Sub
            End If
         End If
      End If
   Next
   
   'check for valid time charges
'   For iList = 0 To LINES_PER_PAGE * MAX_PAGES - 1
   For iList = 0 To UBound(vTimeCard)
      If vTimeCard(iList, TC_HOURS) = "* Error *" Then
         MsgBox "There Is An Invalid Time Entry.", vbExclamation, Caption
         Exit Sub
      End If
   Next
   
   'Update database
   cmdUpdate.Enabled = False
   MouseCursor CURSOR_HOURGLASS
   prg1.Visible = True
   If sOldCard = "" Then
      sNewCard = GetNewNumber()
   Else
      sNewCard = sOldCard
   End If
   bAddingCard = 1
   clsADOCon.BeginTrans
   
   bTimeOrg = False
   If sNewCard = sOldCard Then
      sSql = "DELETE FROM TcitTable WHERE TCCARD='" & sOldCard & "'"
      clsADOCon.ExecuteSql sSql ' rdExecDirect
   Else
      sSql = "INSERT INTO TchdTable (TMCARD,TMEMP,TMDATE,TMDAY," _
          & "TMWEEK) " _
          & "VALUES('" & sNewCard & "'," & Val(cmbEmp) & ",'" _
          & Format(ES_SYSDATE, "mm/dd/yy") & "','" & txtDte & "'" _
          & ",'" & lblWen & "'" & ")"
      clsADOCon.ExecuteSql sSql ' rdExecDirect
   
      bTimeOrg = True
   End If
   
   'process time charges
   'if entering elapsed time, keep track of next timecharge start time
   Dim elapsedStartTime As Variant, ampm As String
   elapsedStartTime = "12:00am"
   ampm = "a"
   
'   For iList = 0 To LINES_PER_PAGE * MAX_PAGES - 1
   For iList = 0 To UBound(vTimeCard)
      If iList <= prg1.Max Then
         Debug.Print CStr(iList)
         prg1.Value = iList
      Else
         Debug.Print 'Progress bar exceeding !00%
      End If
      If Val(vTimeCard(iList, TC_HOURS)) > 0 And Val(vTimeCard(iList, TC_DELETE)) = 0 Then
      
         'if elapsed time, make dummy start and stop times
         If Me.chkElapsed.Value = 1 Then
            vTimeCard(iList, TC_BEGIN) = Replace(Replace(Format(elapsedStartTime, "hh:mm am/pm"), " ", ""), "m", "")
            elapsedStartTime = DateAdd("n", CCur(vTimeCard(iList, TC_HOURS) * 60), elapsedStartTime)
            vTimeCard(iList, TC_END) = Replace(Replace(Format(elapsedStartTime, "hh:mm am/pm"), " ", ""), "m", "")
         End If
      
         sBegTime = vTimeCard(iList, TC_BEGIN)
         'Get the Workcenter Rate,if any.
         'Multiply it by the Regular time rate.
         'Not all companies may want to do it this
         'way and require a setting
         sDebitAcct = ""
         sWorkCenter = ""
         sShop = ""
         sPartNumber = ""
         iRunno = 0
         iOpno = 0
         typeCode = CStr(vTimeCard(iList, TC_TYPE))
         
         If Len(Trim(vTimeCard(iList, TC_MO))) Then
            sPartNumber = vTimeCard(iList, TC_MO)
            sPartNumber = Compress(sPartNumber)
            iRunno = Val(vTimeCard(iList, TC_RUN))
            iOpno = Val(vTimeCard(iList, TC_OP))
         Else
            sDebitAcct = Trim(vTimeCard(iList, TC_ACCT))    'use indirect account selected
         End If
         
         sEndTime = vTimeCard(iList, TC_END)
         iRef = 0
         
         ' Get the start and end time and date from the TCstart and TCend dates.
         ' MM 10/6/2011
'         Dim StartDateTime As Variant, EndDateTime As Variant
'         StartDateTime = txtDte & " " & vTimeCard(iList, TC_BEGIN) & "m"
'         EndDateTime = txtDte & " " & vTimeCard(iList, TC_END) & "m"
'
'
'         If DateDiff("n", StartDateTime, EndDateTime) <= 0 Then
'            EndDateTime = DateAdd("d", 1, EndDateTime)
'         End If
         
         ' get the date
         Dim strStartDate As String
         Dim strEndDate As String
         Dim StartDateTime As Variant
         Dim EndDateTime As Variant
         Dim dtTCStart As Date
         Dim dtTCEnd As Date
         Dim strTCBeg As String
         Dim strTCEnd As String
         
         strStartDate = Format(vTimeCard(iList, TC_STARTTIME), "mm/dd/yy")
         strEndDate = Format(vTimeCard(iList, TC_ENDTIME), "mm/dd/yy")
         strTCBeg = vTimeCard(iList, TC_BEGIN) & "m"
         strTCEnd = vTimeCard(iList, TC_END) & "m"
         
         If (Trim(strStartDate) <> "" And Trim(strEndDate) <> "") Then
            
            ' If the enddate is greater and the starttime is
            ' less than end time then set the enddate same as startdate
            ' 10/25/2011 11:00a ==> 10/26/2011 11:30a
            If ((CDate(strEndDate) > CDate(strStartDate)) And _
                  (CDate(strTCBeg) <= CDate(strTCEnd))) Then
                  strEndDate = Format(vTimeCard(iList, TC_STARTTIME), "mm/dd/yy")
            End If
            
            StartDateTime = strStartDate & " " & vTimeCard(iList, TC_BEGIN) & "m"
            EndDateTime = strEndDate & " " & vTimeCard(iList, TC_END) & "m"
            
         Else
            tc.GetTCardDates CDbl(cmbEmp), CStr(txtDte), _
                           vTimeCard(iList, TC_BEGIN), vTimeCard(iList, TC_END), _
                           dtTCStart, dtTCEnd
            
            StartDateTime = Format(dtTCStart, "mm/dd/yy") & " " & " " & vTimeCard(iList, TC_BEGIN) & "m"
            EndDateTime = Format(dtTCEnd, "mm/dd/yy") & " " & vTimeCard(iList, TC_END) & "m"
            
'            StartDateTime = txtDte & " " & vTimeCard(iList, TC_BEGIN) & "m"
'            EndDateTime = txtDte & " " & vTimeCard(iList, TC_END) & "m"
'
'            If DateDiff("n", StartDateTime, EndDateTime) <= 0 Then
'               EndDateTime = DateAdd("d", 1, EndDateTime)
'            End If
         End If
         
         If tc.CreateTimeCharge(sNewCard, cmbEmp, _
            StartDateTime, EndDateTime, CStr(vTimeCard(iList, TC_TYPE)), _
            CStr(vTimeCard(iList, TC_SRI)), sDebitAcct, _
            sPartNumber, CLng(iRunno), iOpno, iRef, "TS", vTimeCard(iList, TC_ACCEPTED), _
            vTimeCard(iList, TC_REJECTED), vTimeCard(iList, TC_SCRAPPED), CStr(vTimeCard(iList, TC_COMMENTS))) Then
            iRows = iRows + 1
         End If
      End If
   Next
   prg1.Value = 90
   If iRows > 0 Then
      'On Error Resume Next
'      If lstItm.ListCount > 0 Then
'         sBegTime = Mid(lstItm.List(0), 7, 6)
'         sEndTime = Mid(lstItm.List(lstItm.ListCount - 1), 13, 6)
'      End If
      
'      If sNewCard <> sOldCard Then
'         sSql = "INSERT INTO TchdTable (TMCARD,TMEMP,TMDATE,TMDAY," _
'             & "TMWEEK,TMSTART,TMSTOP,TMREGHRS,TMOVTHRS,TMDBLHRS) " _
'             & "VALUES('" & sNewCard & "'," & Val(cmbEmp) & ",'" _
'             & Format(ES_SYSDATE, "mm/dd/yy") & "','" & txtDte & "'" _
'             & ",'" & lblWen & "','" & sBegTime & "','" & sEndTime & "'" _
'             & "," & cRegTime & "," & cOvrTime & "," & cDblTime & ")"
'         clsADOCon.ExecuteSQL sSql ' rdExecDirect
'      End If
      
      tc.ComputeOverlappingCharges cmbEmp, txtDte
            
      tc.UpdateTimeCardTotals cmbEmp, txtDte, bTimeOrg
      
      clsADOCon.CommitTrans
      cmdUpdate.Enabled = True
      prg1.Value = 100
      MouseCursor CURSOR_NORMAL
      SysMsg "The Time Card Was Updated.", True
   Else
      'On Error Resume Next
      clsADOCon.RollbackTrans
      cmdUpdate.Enabled = True
      prg1.Value = 90
      MouseCursor CURSOR_NORMAL
      MsgBox "Couldn't Update The Time Card.", vbExclamation, Caption
   End If
   
   prg1.Visible = False
   bAddingCard = 0
   ResetBoxes
   cmdOps.Visible = False
   optInd.Enabled = True
   'optOhr.Enabled = True
   cmbEmp.Enabled = True
   txtDte.Enabled = True
   cmdCan.Enabled = True
   On Error Resume Next
   cmbEmp.SetFocus
   MouseCursor CURSOR_NORMAL
   Exit Sub
   
DiaErr1:
   sProcName = "UpdateTimeCard"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   On Error Resume Next
   clsADOCon.RollbackTrans
   DoModuleErrors Me
   MouseCursor CURSOR_NORMAL
   
End Sub

Private Function GetRun(iIndex As Integer) As Byte
   
   Dim RdoRun As ADODB.Recordset
   Dim sPartNumber
   sPartNumber = Compress(txtMon(iIndex))
   
   On Error GoTo DiaErr1
   'RdoQry2(0) = sPartNumber
   'RdoQry2(1) = Val(txtRun(iIndex))
   
    cmdObj2.Parameters(0).Value = sPartNumber
    cmdObj2.Parameters(1).Value = Val(txtRun(iIndex))

   bSqlRows = clsADOCon.GetQuerySet(RdoRun, cmdObj2)
   If bSqlRows Then
      With RdoRun
         txtMon(iIndex).ForeColor = Es_TextForeColor
         txtMon(iIndex) = "" & Trim(!PartNum)
         .Cancel
      End With
      GetRun = 1
   Else
      GetRun = 0
      If Len(txtMon(iIndex)) Then
         txtMon(iIndex).ForeColor = ES_RED
      Else
         txtMon(iIndex).ForeColor = Es_TextForeColor
      End If
   End If
   Set RdoRun = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

'Private Sub FillTimeRuns(sFind As String)
'    'fill list box with open runs with matching partial part numbers
'
'    Dim RdoRns As ADODB.Recordset
'    Dim iList  As Integer
'    Dim sPartNumber As String
'
'    lstRns.Clear
'    On Error GoTo DiaErr1
'    sPartNumber = Compress(txtRns)
'    If Len(sPartNumber) > 0 Then
'        sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,PARTREF,PARTNUM " _
'        & "FROM RunsTable,PartTable WHERE RUNREF=PARTREF AND RUNREF Like '" & sPartNumber & "%' " _
'        & "and RUNSTATUS not in ( 'CL', 'CA' ) "
'        bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns)
'            If bSqlRows Then
'                With RdoRns
'                    lstRns.Visible = True
'                    lstRns.Height = 1425
'                    Do Until .EOF
'                        iList = iList + 1
'                        If iList > 300 Then Exit Do
'                        lstRns.AddItem "" & !PARTNUM & Chr(9) & Format(!RunNo, "####0")
'                        .MoveNext
'                    Loop
'                    .Cancel
'                    On Error Resume Next
'                    lstRns.SetFocus
'                End With
'            End If
'    Else
'        MsgBox "Enter At Least (1) Leading Character.", _
'            vbInformation, Caption
'    End If
'    DoEvents
'    On Error Resume Next
'    Set RdoRns = Nothing
'    Exit Sub
'
'DiaErr1:
'    sProcName = "filltimeruns"
'    CurrError.Number = Err.Number
'    CurrError.Description = Err.Description
'    DoModuleErrors Me
'
'End Sub
'
'
'Private Sub FillAccounts(sFind As String)
'    'fill list box with matching partial accounts for indirect time charges
'
'    Dim RdoRns As ADODB.Recordset
'    Dim iList  As Integer
'    Dim sNumber As String
'
'    lstRns.Clear
'    On Error GoTo DiaErr1
'    sNumber = Compress(txtRns)
'    If Len(sNumber) > 0 Then
'        sSql = "select GLACCTNO, GLDESCR from GlacTable " _
'        & "where GLACCTNO Like '" & sNumber & "%' " _
'        & "and GLACCTREF not in (select distinct GLMASTER from GlacTable)"
'        bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns)
'            If bSqlRows Then
'                With RdoRns
'                    lstRns.Visible = True
'                    lstRns.Height = 1425
'                    Do Until .EOF
'                        iList = iList + 1
'                        If iList > 300 Then Exit Do
'                        lstRns.AddItem "" & !GLACCTNO & Chr(9) & !GLDESCR
'                        .MoveNext
'                    Loop
'                    .Cancel
'                    On Error Resume Next
'                    lstRns.SetFocus
'                End With
'            End If
'    Else
'        MsgBox "Enter At Least (1) Leading Character.", _
'            vbInformation, Caption
'    End If
'    DoEvents
'    On Error Resume Next
'    Set RdoRns = Nothing
'    Exit Sub
'
'DiaErr1:
'    sProcName = "fillaccounts"
'    CurrError.Number = Err.Number
'    CurrError.Description = Err.Description
'    DoModuleErrors Me
'
'End Sub
'

Private Sub FillTimeRuns(sFind As String)
   'fill list box with open runs with matching partial part numbers
   
   Dim RdoRns As ADODB.Recordset
   Dim sPart As String
   
   cboFind.Clear
   cboFind.Visible = True
   On Error GoTo DiaErr1
   
   sSql = "SELECT RUNNO, PARTNUM " _
          & "from RunsTable join PartTable on RUNREF = PARTREF AND RUNREF Like '" & sFind & "%' " _
          & "and RUNSTATUS not in ( 'CL', 'CA' ) "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            sPart = Trim(!PartNum)
            cboFind.AddItem "" & sPart & Space(31 - Len(sPart)) & Format(!Runno, "####0")
            .MoveNext
         Loop
         .Cancel
         On Error Resume Next
         cboFind.SetFocus
      End With
   End If
   
   If cboFind.ListCount > 0 Then
      cboFind.ListIndex = 0
   End If
   
   'DoEvents
   On Error Resume Next
   Set RdoRns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "filltimeruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FillAccounts(sFind As String)
   'fill list box with matching partial accounts for indirect time charges
   
   Dim RdoRns As ADODB.Recordset
   Dim sAcct As String
   
   cboFind.Clear
   cboFind.Visible = True
   On Error GoTo DiaErr1
   sSql = "select GLACCTNO, GLDESCR from GlacTable " _
          & "where GLACCTNO Like '" & sFind & "%' " _
          & "and GLACCTREF not in (select distinct GLMASTER from GlacTable)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            sAcct = Trim(!GLACCTNO)
            cboFind.AddItem "" & sAcct & Space(13 - Len(sAcct)) & Trim(!GLDESCR)
            .MoveNext
         Loop
         .Cancel
         On Error Resume Next
         cboFind.SetFocus
      End With
   End If
   
   If cboFind.ListCount > 0 Then
      cboFind.ListIndex = 0
   End If
   
   'DoEvents
   On Error Resume Next
   Set RdoRns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillaccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub GetWeekEnd()
   Dim RdoGet As ADODB.Recordset
   Dim A As Integer
   Dim iList As Integer
   Dim dDate As Date
   Dim sWeekEnds As String
   
   On Error GoTo DiaErr1
   sSql = "SELECT WEEKENDS FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         sWeekEnds = "" & Trim(!WEEKENDS)
         .Cancel
      End With
      If sWeekEnds = "Sat" Then iList = 7 Else iList = 8
   End If
   dDate = txtDte
   A = Format(txtDte, "w")
   lblWen = Format(dDate + (iList - A), "mm/dd/yy")
   Set RdoGet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getweeken"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtRun_LostFocus(Index As Integer)
   'test for invalid part/run/op
   bGood(Index + iIndex) = TestRunOp(Index)
   vTimeCard(Index + iIndex, TC_RUN) = txtRun(Index)
   
   If Not bGood(Index + iIndex) Then
      Exit Sub
   End If
   
   'vTimeCard(Index + iIndex, TC_RUN) = txtRun(Index)
   
End Sub

Private Sub txtSri_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtSri_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtSri_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub txtSri_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub txtSri_LostFocus(Index As Integer)
   txtSri(Index) = CheckLen(txtSri(Index), 1)
   If optTyp(Index).Value = vbUnchecked Then
      txtSri(Index) = "I"
   Else
      If txtSri(Index) <> "S" And txtSri(Index) <> "R" And txtSri(Index) <> "T" Then
         txtSri(Index) = "R"
      End If
   End If
   vTimeCard(Index + iIndex, TC_SRI) = txtSri(Index)
   
End Sub

Private Sub GetRunOp(Index2 As Integer)
   Dim RdoGet As ADODB.Recordset
   Dim sPartNumber As String
   sPartNumber = txtMon(Index2)
   sSql = "SELECT OPREF,OPRUN,OPNO,OPCENTER,WCNREF,WCNACCT FROM RnopTable," _
          & "WcntTable WHERE (OPCENTER=WCNREF AND OPREF='" & sPartNumber & "' AND OPRUN=" _
          & Val(txtRun(Index2)) & " AND OPNO=" & Val(txtOpn(Index2)) & ") "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet)
   If bSqlRows Then
      With RdoGet
         If Trim(!WCNACCT) <> "" Then
            vTimeCard(iIndex + Index2, TC_TIMEACCT) = "" & Trim(!WCNACCT)
         ElseIf sEmplAcct <> "" Then _
               vTimeCard(iIndex + Index2, TC_TIMEACCT) = sEmplAcct
         Else
            vTimeCard(iIndex + Index2, TC_TIMEACCT) = sCoTimeAcct
         End If
         'Debug.Print "GetRunOp index " & Index2 & " OP " & txtOpn(Index2).Text & " = " & Trim(!OPCENTER)
         vTimeCard(iIndex + Index2, TC_WC) = "" & Trim(!OPCENTER)
         .Cancel
      End With
   Else
      If sEmplAcct <> "" Then
         vTimeCard(iIndex + Index2, TC_TIMEACCT) = sEmplAcct
      Else
         vTimeCard(iIndex + Index2, TC_TIMEACCT) = sCoTimeAcct
      End If
      vTimeCard(iIndex + Index2, TC_WC) = sEmplCenter
   End If
   Set RdoGet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getrunop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetNewNumber() As String
   Dim S As Single
   Dim l As Long
   Dim m As Long
   Dim t As String
   On Error Resume Next
   '    m = DateValue(Format(ES_SYSDATE, "yyyy,mm,dd"))
   '    s = TimeValue(Format(ES_SYSDATE, "hh:mm:ss"))
   '    l = s * 1000000
   '    GetNewNumber = Format(m, "00000") & Format(l, "000000")
   Dim dt As Variant
   dt = GetServerDateTime()
   m = DateValue(Format(dt, "yyyy,mm,dd"))
   S = TimeValue(Format(dt, "hh:mm:ss"))
   l = S * 1000000
   GetNewNumber = Format(m, "00000") & Format(l, "000000")
   
End Function

Private Sub GetCardTime()
   Dim iList As Integer
   Dim cTotal As Currency
'   For iList = 0 To LINES_PER_PAGE * MAX_PAGES - 1
   For iList = 0 To UBound(vTimeCard)
      cTotal = cTotal + Val(vTimeCard(iList, TC_HOURS))
   Next
   lblTot = Format(Round(cTotal, 2), "###0.000")
   
End Sub

Private Function TestRunOp(Index As Integer) As Boolean
   'test part#/run#/op# for a valid combination
   
   'if one of the fields is blank, just return false.  Data entry is still going on
   TestRunOp = False
   If Len(Trim(txtMon(Index).Text)) = 0 Or Len(Trim(txtRun(Index).Text)) = 0 _
          Or Len(Trim(txtOpn(Index).Text)) = 0 Then
      Exit Function
   End If
   
   'there is something in all 3 fields.  It should be valid
   bGood(Index + iIndex) = IsValidRunOp(txtMon(Index), txtRun(Index), txtOpn(Index), False, False)
   Dim color As Long
   If bGood(Index + iIndex) Then
      color = Es_TextForeColor
      TestRunOp = True
   Else
      color = ES_RED
   End If
   txtMon(Index).ForeColor = color
   txtRun(Index).ForeColor = color
   txtOpn(Index).ForeColor = color
   
   
End Function

Private Function TestAccount(Index As Integer) As Boolean
   'test for valid account
   bGood(Index + iIndex) = IsValidAccount(txtMon(Index))
   If bGood(Index + iIndex) Then
      txtMon(Index).ForeColor = Es_TextForeColor
      TestAccount = True
   Else
      txtMon(Index).ForeColor = ES_RED
      TestAccount = False
   End If
   
End Function

Private Sub TurnElapsedTimeOnOrOff()
   Dim I As Integer
   For I = 0 To LINES_PER_PAGE - 1
      txtHrs(I).Enabled = chkElapsed.Value
      txtBeg(I).Enabled = 1 - chkElapsed.Value
      txtEnd(I).Enabled = 1 - chkElapsed.Value
   Next
End Sub

Private Sub ThereIsAnError(Index As Integer)
   lblErrors(Index).ForeColor = ES_RED
   lblErrors(Index) = "* Error *"
End Sub

Private Sub ThereIsNoError(Index As Integer)
   lblErrors(Index).ForeColor = Es_TextForeColor
   lblErrors(Index) = ""
End Sub

Sub ChangeRow(NewRow As Integer)
   If NewRow <> iCurrIndex Then
      iCurrIndex = NewRow
      cboFind.Clear
      cboFind.Visible = False
   End If
End Sub


Private Sub ClearArray()
Dim iList As Integer

   'set time card to defaults
'   For iList = 0 To LINES_PER_PAGE * MAX_PAGES - 1
    For iList = 0 To UBound(vTimeCard, 1)
      vTimeCard(iList, TC_DI) = 1
      vTimeCard(iList, TC_MO) = ""
      vTimeCard(iList, TC_ACCT) = ""
      vTimeCard(iList, TC_RUN) = ""
      vTimeCard(iList, TC_OP) = ""
      vTimeCard(iList, TC_TYPE) = ""
      vTimeCard(iList, TC_BEGIN) = ""
      vTimeCard(iList, TC_END) = ""
      vTimeCard(iList, TC_HOURS) = ""
      vTimeCard(iList, TC_SRI) = ""
      vTimeCard(iList, TC_WC) = ""
      vTimeCard(iList, TC_TIMEACCT) = ""
      vTimeCard(iList, TC_DELETE) = 0
      vTimeCard(iList, TC_ACCEPTED) = 0
      vTimeCard(iList, TC_REJECTED) = 0
      vTimeCard(iList, TC_SCRAPPED) = 0
      vTimeCard(iList, TC_COMMENTS) = ""
   Next
End Sub
