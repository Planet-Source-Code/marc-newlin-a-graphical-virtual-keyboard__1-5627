VERSION 5.00
Begin VB.Form vkb 
   BackColor       =   &H80000007&
   Caption         =   "Virtual Keyboard"
   ClientHeight    =   3930
   ClientLeft      =   585
   ClientTop       =   3705
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   13875
   Begin VB.CommandButton DEL 
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   1560
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   9360
      Top             =   2400
   End
   Begin VB.CommandButton ENTER2 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton ADDI 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton NUMPAD9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton NUMPAD0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton DEC 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton NUMPAD4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton NUMPAD5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton NUMPAD6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton NUMPAD1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton NUMPAD2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton NUMPAD3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton NUMPAD7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton NUM 
      Caption         =   "Num Lock"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton DIV 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton NUMPAD8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton SUBT 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton MULT 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton UP 
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton RIGHTB 
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton DOWN 
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton LEFTB 
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton PGDOWN 
      Caption         =   "Page Down"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton END 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton PGUP 
      Caption         =   "Page Up"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton HOME 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton INSERT 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton SPACEBAR 
      Caption         =   "Space Bar"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton CONTROL2 
      Caption         =   "Ctrl"
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton OPT 
      Height          =   495
      Left            =   7440
      Picture         =   "vkb.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton START2 
      Height          =   495
      Left            =   6600
      Picture         =   "vkb.frx":0556
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton ALT2 
      Caption         =   "Alt"
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton ALT1 
      Caption         =   "Alt"
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton START1 
      Height          =   495
      Left            =   960
      MaskColor       =   &H00FFFFFF&
      Picture         =   "vkb.frx":0AA4
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton CONTROL1 
      Caption         =   "Ctrl"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton OEM_2 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton ESCAPE 
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton X 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton C 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton V 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton B 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton N 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton M 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton OEM_COMMA 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Z 
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton SHIFT2 
      Caption         =   "Shift"
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton SHIFT1 
      Caption         =   "Shift"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton K 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton J 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton L 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton H 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton G 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton F 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton D 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton S 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton A 
      BackColor       =   &H8000000A&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton OEM_1 
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton OEM_7 
      Caption         =   "'"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton CAPS 
      Caption         =   "Caps Lock"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton ENTER 
      Caption         =   "Enter"
      Height          =   1095
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton OEM_5 
      Caption         =   "\"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton OEM_6 
      Caption         =   "]"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton OEM_4 
      Caption         =   "["
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton P 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton O 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton I 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton U 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Y 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton T 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton R 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton E 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton W 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Q 
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton TAB 
      Caption         =   "Tab"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Back 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton OEM_PLUS 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton OEM_MINUS 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton N0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton N9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton N8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton N7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton N6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton N5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton N4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton N3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton N2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton N1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton OEM_3 
      Caption         =   "`"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton PAUSEKEY 
      Caption         =   "Pause "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton PRINTSCR 
      Caption         =   "Print Scrn"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton SCROLL 
      Caption         =   "Scroll Lock"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton F1 
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton F12 
      Caption         =   "F12"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton F11 
      Caption         =   "F11"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton F9 
      Caption         =   "F9"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton F2 
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton F3 
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton F4 
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton F10 
      Caption         =   "F10"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton F5 
      Caption         =   "F5"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton F6 
      Caption         =   "F6"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton F7 
      Caption         =   "F7"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton F8 
      Caption         =   "F8"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton OEM_PERIOD 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   11400
      Picture         =   "vkb.frx":0FF2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2355
   End
End
Attribute VB_Name = "vkb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
CheckKeys
End Sub


