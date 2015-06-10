VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ANNO 1404 - Produktketten Rechner"
   ClientHeight    =   8865
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11355
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   59
      Left            =   15480
      Picture         =   "MainForm.frx":35D33
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   285
      Top             =   8160
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   58
      Left            =   14520
      Picture         =   "MainForm.frx":36B09
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   284
      Top             =   8160
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sonstige"
      Height          =   375
      Left            =   2760
      TabIndex        =   152
      Top             =   1200
      Width           =   1215
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   57
      Left            =   13560
      Picture         =   "MainForm.frx":378B4
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   150
      Top             =   8160
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   56
      Left            =   12600
      Picture         =   "MainForm.frx":384C7
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   149
      Top             =   8280
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   55
      Left            =   11640
      Picture         =   "MainForm.frx":39213
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   148
      Top             =   8280
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   54
      Left            =   16440
      Picture         =   "MainForm.frx":3A02F
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   147
      Top             =   7320
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   53
      Left            =   15480
      Picture         =   "MainForm.frx":3ADD2
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   146
      Top             =   7320
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   52
      Left            =   14520
      Picture         =   "MainForm.frx":3BB69
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   145
      Top             =   7320
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   51
      Left            =   13560
      Picture         =   "MainForm.frx":3C9BB
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   144
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Orient"
      Height          =   375
      Left            =   1440
      TabIndex        =   30
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Okzident"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gesamtkosten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   119
      Top             =   6240
      Width           =   3855
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   1320
         Picture         =   "MainForm.frx":3D5DC
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   132
         ToolTipText     =   "Werkzeug"
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   1320
         Picture         =   "MainForm.frx":3D9FB
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   131
         ToolTipText     =   "Stein"
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   120
         Picture         =   "MainForm.frx":3DE45
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   124
         ToolTipText     =   "Goldmünzen"
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   2520
         Picture         =   "MainForm.frx":3E284
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   123
         ToolTipText     =   "Glas"
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   120
         Picture         =   "MainForm.frx":3E6FF
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   121
         ToolTipText     =   "Holz"
         Top             =   720
         Width           =   375
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   375
            Index           =   7
            Left            =   360
            TabIndex        =   122
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   2520
         Picture         =   "MainForm.frx":3EB07
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   120
         ToolTipText     =   "Mosaik"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label gesamt_kosten 
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   134
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Betriebskosten:"
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
         Left            =   120
         TabIndex        =   133
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label gesamt_geld 
         Caption         =   "0"
         Height          =   255
         Left            =   480
         TabIndex        =   130
         Top             =   360
         Width           =   735
      End
      Begin VB.Label gesamt_werkzeug 
         Caption         =   "0"
         Height          =   255
         Left            =   1680
         TabIndex        =   129
         Top             =   360
         Width           =   735
      End
      Begin VB.Label gesamt_glas 
         Caption         =   "0"
         Height          =   255
         Left            =   2880
         TabIndex        =   128
         Top             =   360
         Width           =   735
      End
      Begin VB.Label gesamt_holz 
         Caption         =   "0"
         Height          =   255
         Left            =   480
         TabIndex        =   127
         Top             =   720
         Width           =   735
      End
      Begin VB.Label gesamt_stein 
         Caption         =   "0"
         Height          =   255
         Left            =   1680
         TabIndex        =   126
         Top             =   720
         Width           =   735
      End
      Begin VB.Label gesamt_mosaik 
         Caption         =   "0"
         Height          =   255
         Left            =   2880
         TabIndex        =   125
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   50
      Left            =   12600
      Picture         =   "MainForm.frx":3EF89
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   118
      Top             =   7320
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   49
      Left            =   11640
      Picture         =   "MainForm.frx":3FBB3
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   117
      Top             =   7320
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   48
      Left            =   16440
      Picture         =   "MainForm.frx":407FB
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   116
      Top             =   6360
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   47
      Left            =   15480
      Picture         =   "MainForm.frx":413E6
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   115
      Top             =   6360
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   46
      Left            =   14520
      Picture         =   "MainForm.frx":41F33
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   114
      Top             =   6360
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   45
      Left            =   13560
      Picture         =   "MainForm.frx":42B3F
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   113
      Top             =   6360
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   44
      Left            =   12600
      Picture         =   "MainForm.frx":43705
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   112
      Top             =   6360
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   43
      Left            =   11640
      Picture         =   "MainForm.frx":44314
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   111
      Top             =   6360
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   42
      Left            =   16440
      Picture         =   "MainForm.frx":44EB4
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   110
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   41
      Left            =   15480
      Picture         =   "MainForm.frx":45AF4
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   109
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   40
      Left            =   14520
      Picture         =   "MainForm.frx":46787
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   108
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   39
      Left            =   13560
      Picture         =   "MainForm.frx":4731C
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   107
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   38
      Left            =   12600
      Picture         =   "MainForm.frx":47EB6
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   106
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   37
      Left            =   11640
      Picture         =   "MainForm.frx":48C7D
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   105
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   36
      Left            =   16440
      Picture         =   "MainForm.frx":497E3
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   104
      Top             =   4440
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   35
      Left            =   15480
      Picture         =   "MainForm.frx":4A597
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   103
      Top             =   4440
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   34
      Left            =   14520
      Picture         =   "MainForm.frx":4B31F
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   102
      Top             =   4440
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   33
      Left            =   13560
      Picture         =   "MainForm.frx":4C0C8
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   101
      Top             =   4440
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   32
      Left            =   12600
      Picture         =   "MainForm.frx":4CE95
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   100
      Top             =   4440
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   31
      Left            =   11640
      Picture         =   "MainForm.frx":4DC20
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   99
      Top             =   4440
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   30
      Left            =   16440
      Picture         =   "MainForm.frx":4E9BD
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   98
      Top             =   3480
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   29
      Left            =   15480
      Picture         =   "MainForm.frx":4F72D
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   97
      Top             =   3480
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   28
      Left            =   14520
      Picture         =   "MainForm.frx":504E5
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   96
      Top             =   3480
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   27
      Left            =   13560
      Picture         =   "MainForm.frx":510ED
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   95
      Top             =   3480
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   26
      Left            =   12600
      Picture         =   "MainForm.frx":51D20
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   94
      Top             =   3480
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   25
      Left            =   11640
      Picture         =   "MainForm.frx":528CB
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   93
      Top             =   3480
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   24
      Left            =   16440
      Picture         =   "MainForm.frx":5347C
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   92
      Top             =   2640
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   23
      Left            =   15480
      Picture         =   "MainForm.frx":5415B
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   91
      Top             =   2640
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   22
      Left            =   14520
      Picture         =   "MainForm.frx":54F13
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   90
      Top             =   2640
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   21
      Left            =   13560
      Picture         =   "MainForm.frx":55D19
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   89
      Top             =   2640
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   20
      Left            =   12600
      Picture         =   "MainForm.frx":56AC4
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   88
      Top             =   2640
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   19
      Left            =   11640
      Picture         =   "MainForm.frx":57691
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   87
      Top             =   2640
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   18
      Left            =   16440
      Picture         =   "MainForm.frx":58457
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   86
      Top             =   1800
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   17
      Left            =   15480
      Picture         =   "MainForm.frx":59180
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   85
      Top             =   1800
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   16
      Left            =   14520
      Picture         =   "MainForm.frx":59D7A
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   84
      Top             =   1800
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   15
      Left            =   13560
      Picture         =   "MainForm.frx":5A9A9
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   83
      Top             =   1800
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   14
      Left            =   12600
      Picture         =   "MainForm.frx":5B5C7
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   82
      Top             =   1800
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   13
      Left            =   11640
      Picture         =   "MainForm.frx":5C360
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   81
      Top             =   1800
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   12
      Left            =   16440
      Picture         =   "MainForm.frx":5D0B2
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   80
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   11
      Left            =   15480
      Picture         =   "MainForm.frx":5DE04
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   79
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   10
      Left            =   14520
      Picture         =   "MainForm.frx":5EBAA
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   78
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   9
      Left            =   13560
      Picture         =   "MainForm.frx":5F8F5
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   77
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   8
      Left            =   12600
      Picture         =   "MainForm.frx":606D3
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   76
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   7
      Left            =   11640
      Picture         =   "MainForm.frx":613E2
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   75
      Top             =   840
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   6
      Left            =   16440
      Picture         =   "MainForm.frx":61FE5
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   74
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   0
      Left            =   9240
      Picture         =   "MainForm.frx":62C1D
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   73
      Top             =   9240
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   11640
      Picture         =   "MainForm.frx":6384E
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   71
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   2
      Left            =   12600
      Picture         =   "MainForm.frx":6457E
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   70
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   3
      Left            =   13560
      Picture         =   "MainForm.frx":65229
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   69
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   4
      Left            =   14520
      Picture         =   "MainForm.frx":65E5A
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   68
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   5
      Left            =   15480
      Picture         =   "MainForm.frx":66A6A
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   67
      Top             =   0
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   46
      Top             =   0
      Width           =   3855
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         Picture         =   "MainForm.frx":6762D
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   47
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Produktionsketten Rechner"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   48
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Berechnen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Top             =   7920
      Width           =   3855
   End
   Begin VB.Frame outFrame 
      Caption         =   "Produktionskette"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   4080
      TabIndex        =   36
      Top             =   0
      Width           =   7215
      Begin VB.Frame out_Frame 
         Height          =   1215
         Index           =   6
         Left            =   120
         TabIndex        =   263
         Top             =   7440
         Width           =   6975
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   6
            Left            =   3360
            Picture         =   "MainForm.frx":68141
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   270
            ToolTipText     =   "Holz"
            Top             =   720
            Width           =   375
            Begin VB.Label Label1 
               Caption         =   "Label1"
               Height          =   375
               Index           =   6
               Left            =   360
               TabIndex        =   271
               Top             =   0
               Width           =   15
            End
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   6
            Left            =   4440
            Picture         =   "MainForm.frx":68549
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   269
            ToolTipText     =   "Werkzeug"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   6
            Left            =   4440
            Picture         =   "MainForm.frx":68968
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   268
            ToolTipText     =   "Stein"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   6
            Left            =   5640
            Picture         =   "MainForm.frx":68DB2
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   267
            ToolTipText     =   "Glas"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   6
            Left            =   3360
            Picture         =   "MainForm.frx":6922D
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   266
            ToolTipText     =   "Goldmünzen"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture7 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   6
            Left            =   5640
            Picture         =   "MainForm.frx":6966C
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   265
            ToolTipText     =   "Mosaik"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox item_pic 
            BorderStyle     =   0  'None
            Height          =   855
            Index           =   6
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   855
            TabIndex        =   264
            Top             =   240
            Width           =   855
         End
         Begin VB.Label out_needed 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   2280
            TabIndex        =   283
            Top             =   360
            Width           =   855
         End
         Begin VB.Label out_geld 
            Caption         =   "0"
            Height          =   255
            Index           =   6
            Left            =   3720
            TabIndex        =   282
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_holz 
            Caption         =   "0"
            Height          =   255
            Index           =   6
            Left            =   3720
            TabIndex        =   281
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_werkzeug 
            Caption         =   "0"
            Height          =   255
            Index           =   6
            Left            =   4920
            TabIndex        =   280
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_stein 
            Caption         =   "0"
            Height          =   255
            Index           =   6
            Left            =   4920
            TabIndex        =   279
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_glas 
            Caption         =   "0"
            Height          =   255
            Index           =   6
            Left            =   6120
            TabIndex        =   278
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_mosaik 
            Caption         =   "0"
            Height          =   255
            Index           =   6
            Left            =   6120
            TabIndex        =   277
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_betrieb 
            Caption         =   "0"
            Height          =   255
            Index           =   6
            Left            =   2280
            TabIndex        =   276
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Betriebskosten:"
            Height          =   255
            Index           =   6
            Left            =   1080
            TabIndex        =   275
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Benötigt:"
            Height          =   255
            Index           =   6
            Left            =   1080
            TabIndex        =   274
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label out_last 
            Caption         =   "0%"
            Height          =   255
            Index           =   6
            Left            =   2280
            TabIndex        =   273
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label22 
            Caption         =   "Auslastung:"
            Height          =   255
            Index           =   6
            Left            =   1080
            TabIndex        =   272
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame out_Frame 
         Height          =   1215
         Index           =   5
         Left            =   120
         TabIndex        =   242
         Top             =   6240
         Width           =   6975
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   5
            Left            =   3360
            Picture         =   "MainForm.frx":69AEE
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   249
            ToolTipText     =   "Holz"
            Top             =   720
            Width           =   375
            Begin VB.Label Label1 
               Caption         =   "Label1"
               Height          =   375
               Index           =   5
               Left            =   360
               TabIndex        =   250
               Top             =   0
               Width           =   15
            End
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   5
            Left            =   4440
            Picture         =   "MainForm.frx":69EF6
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   248
            ToolTipText     =   "Werkzeug"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   5
            Left            =   4440
            Picture         =   "MainForm.frx":6A315
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   247
            ToolTipText     =   "Stein"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   5
            Left            =   5640
            Picture         =   "MainForm.frx":6A75F
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   246
            ToolTipText     =   "Glas"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   5
            Left            =   3360
            Picture         =   "MainForm.frx":6ABDA
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   245
            ToolTipText     =   "Goldmünzen"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture7 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   5
            Left            =   5640
            Picture         =   "MainForm.frx":6B019
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   244
            ToolTipText     =   "Mosaik"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox item_pic 
            BorderStyle     =   0  'None
            Height          =   855
            Index           =   5
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   855
            TabIndex        =   243
            Top             =   240
            Width           =   855
         End
         Begin VB.Label out_needed 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   2280
            TabIndex        =   262
            Top             =   360
            Width           =   855
         End
         Begin VB.Label out_geld 
            Caption         =   "0"
            Height          =   255
            Index           =   5
            Left            =   3720
            TabIndex        =   261
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_holz 
            Caption         =   "0"
            Height          =   255
            Index           =   5
            Left            =   3720
            TabIndex        =   260
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_werkzeug 
            Caption         =   "0"
            Height          =   255
            Index           =   5
            Left            =   4920
            TabIndex        =   259
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_stein 
            Caption         =   "0"
            Height          =   255
            Index           =   5
            Left            =   4920
            TabIndex        =   258
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_glas 
            Caption         =   "0"
            Height          =   255
            Index           =   5
            Left            =   6120
            TabIndex        =   257
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_mosaik 
            Caption         =   "0"
            Height          =   255
            Index           =   5
            Left            =   6120
            TabIndex        =   256
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_betrieb 
            Caption         =   "0"
            Height          =   255
            Index           =   5
            Left            =   2280
            TabIndex        =   255
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Betriebskosten:"
            Height          =   255
            Index           =   5
            Left            =   1080
            TabIndex        =   254
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Benötigt:"
            Height          =   255
            Index           =   5
            Left            =   1080
            TabIndex        =   253
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label out_last 
            Caption         =   "0%"
            Height          =   255
            Index           =   5
            Left            =   2280
            TabIndex        =   252
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label22 
            Caption         =   "Auslastung:"
            Height          =   255
            Index           =   5
            Left            =   1080
            TabIndex        =   251
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame out_Frame 
         Height          =   1215
         Index           =   4
         Left            =   120
         TabIndex        =   221
         Top             =   5040
         Width           =   6975
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   4
            Left            =   3360
            Picture         =   "MainForm.frx":6B49B
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   228
            ToolTipText     =   "Holz"
            Top             =   720
            Width           =   375
            Begin VB.Label Label1 
               Caption         =   "Label1"
               Height          =   375
               Index           =   4
               Left            =   360
               TabIndex        =   229
               Top             =   0
               Width           =   15
            End
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   4
            Left            =   4440
            Picture         =   "MainForm.frx":6B8A3
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   227
            ToolTipText     =   "Werkzeug"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   4
            Left            =   4440
            Picture         =   "MainForm.frx":6BCC2
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   226
            ToolTipText     =   "Stein"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   4
            Left            =   5640
            Picture         =   "MainForm.frx":6C10C
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   225
            ToolTipText     =   "Glas"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   4
            Left            =   3360
            Picture         =   "MainForm.frx":6C587
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   224
            ToolTipText     =   "Goldmünzen"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture7 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   4
            Left            =   5640
            Picture         =   "MainForm.frx":6C9C6
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   223
            ToolTipText     =   "Mosaik"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox item_pic 
            BorderStyle     =   0  'None
            Height          =   855
            Index           =   4
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   855
            TabIndex        =   222
            Top             =   240
            Width           =   855
         End
         Begin VB.Label out_needed 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   2280
            TabIndex        =   241
            Top             =   360
            Width           =   855
         End
         Begin VB.Label out_geld 
            Caption         =   "0"
            Height          =   255
            Index           =   4
            Left            =   3720
            TabIndex        =   240
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_holz 
            Caption         =   "0"
            Height          =   255
            Index           =   4
            Left            =   3720
            TabIndex        =   239
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_werkzeug 
            Caption         =   "0"
            Height          =   255
            Index           =   4
            Left            =   4920
            TabIndex        =   238
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_stein 
            Caption         =   "0"
            Height          =   255
            Index           =   4
            Left            =   4920
            TabIndex        =   237
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_glas 
            Caption         =   "0"
            Height          =   255
            Index           =   4
            Left            =   6120
            TabIndex        =   236
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_mosaik 
            Caption         =   "0"
            Height          =   255
            Index           =   4
            Left            =   6120
            TabIndex        =   235
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_betrieb 
            Caption         =   "0"
            Height          =   255
            Index           =   4
            Left            =   2280
            TabIndex        =   234
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Betriebskosten:"
            Height          =   255
            Index           =   4
            Left            =   1080
            TabIndex        =   233
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Benötigt:"
            Height          =   255
            Index           =   4
            Left            =   1080
            TabIndex        =   232
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label out_last 
            Caption         =   "0%"
            Height          =   255
            Index           =   4
            Left            =   2280
            TabIndex        =   231
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label22 
            Caption         =   "Auslastung:"
            Height          =   255
            Index           =   4
            Left            =   1080
            TabIndex        =   230
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame out_Frame 
         Height          =   1215
         Index           =   3
         Left            =   120
         TabIndex        =   200
         Top             =   3840
         Width           =   6975
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   3
            Left            =   3360
            Picture         =   "MainForm.frx":6CE48
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   207
            ToolTipText     =   "Holz"
            Top             =   720
            Width           =   375
            Begin VB.Label Label1 
               Caption         =   "Label1"
               Height          =   375
               Index           =   3
               Left            =   360
               TabIndex        =   208
               Top             =   0
               Width           =   15
            End
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   3
            Left            =   4440
            Picture         =   "MainForm.frx":6D250
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   206
            ToolTipText     =   "Werkzeug"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   3
            Left            =   4440
            Picture         =   "MainForm.frx":6D66F
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   205
            ToolTipText     =   "Stein"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   3
            Left            =   5640
            Picture         =   "MainForm.frx":6DAB9
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   204
            ToolTipText     =   "Glas"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   3
            Left            =   3360
            Picture         =   "MainForm.frx":6DF34
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   203
            ToolTipText     =   "Goldmünzen"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture7 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   3
            Left            =   5640
            Picture         =   "MainForm.frx":6E373
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   202
            ToolTipText     =   "Mosaik"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox item_pic 
            BorderStyle     =   0  'None
            Height          =   855
            Index           =   3
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   855
            TabIndex        =   201
            Top             =   240
            Width           =   855
         End
         Begin VB.Label out_needed 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   220
            Top             =   360
            Width           =   855
         End
         Begin VB.Label out_geld 
            Caption         =   "0"
            Height          =   255
            Index           =   3
            Left            =   3720
            TabIndex        =   219
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_holz 
            Caption         =   "0"
            Height          =   255
            Index           =   3
            Left            =   3720
            TabIndex        =   218
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_werkzeug 
            Caption         =   "0"
            Height          =   255
            Index           =   3
            Left            =   4920
            TabIndex        =   217
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_stein 
            Caption         =   "0"
            Height          =   255
            Index           =   3
            Left            =   4920
            TabIndex        =   216
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_glas 
            Caption         =   "0"
            Height          =   255
            Index           =   3
            Left            =   6120
            TabIndex        =   215
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_mosaik 
            Caption         =   "0"
            Height          =   255
            Index           =   3
            Left            =   6120
            TabIndex        =   214
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_betrieb 
            Caption         =   "0"
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   213
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Betriebskosten:"
            Height          =   255
            Index           =   3
            Left            =   1080
            TabIndex        =   212
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Benötigt:"
            Height          =   255
            Index           =   3
            Left            =   1080
            TabIndex        =   211
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label out_last 
            Caption         =   "0%"
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   210
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label22 
            Caption         =   "Auslastung:"
            Height          =   255
            Index           =   3
            Left            =   1080
            TabIndex        =   209
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame out_Frame 
         Height          =   1215
         Index           =   2
         Left            =   120
         TabIndex        =   179
         Top             =   2640
         Width           =   6975
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   3360
            Picture         =   "MainForm.frx":6E7F5
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   186
            ToolTipText     =   "Holz"
            Top             =   720
            Width           =   375
            Begin VB.Label Label1 
               Caption         =   "Label1"
               Height          =   375
               Index           =   2
               Left            =   360
               TabIndex        =   187
               Top             =   0
               Width           =   15
            End
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   4440
            Picture         =   "MainForm.frx":6EBFD
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   185
            ToolTipText     =   "Werkzeug"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   4440
            Picture         =   "MainForm.frx":6F01C
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   184
            ToolTipText     =   "Stein"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   5640
            Picture         =   "MainForm.frx":6F466
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   183
            ToolTipText     =   "Glas"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   3360
            Picture         =   "MainForm.frx":6F8E1
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   182
            ToolTipText     =   "Goldmünzen"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture7 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   5640
            Picture         =   "MainForm.frx":6FD20
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   181
            ToolTipText     =   "Mosaik"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox item_pic 
            BorderStyle     =   0  'None
            Height          =   855
            Index           =   2
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   855
            TabIndex        =   180
            Top             =   240
            Width           =   855
         End
         Begin VB.Label out_needed 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   199
            Top             =   360
            Width           =   855
         End
         Begin VB.Label out_geld 
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   198
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_holz 
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   197
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_werkzeug 
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   4920
            TabIndex        =   196
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_stein 
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   4920
            TabIndex        =   195
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_glas 
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   6120
            TabIndex        =   194
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_mosaik 
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   6120
            TabIndex        =   193
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_betrieb 
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   192
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Betriebskosten:"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   191
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Benötigt:"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   190
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label out_last 
            Caption         =   "0%"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   189
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label22 
            Caption         =   "Auslastung:"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   188
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame out_Frame 
         Height          =   1215
         Index           =   1
         Left            =   120
         TabIndex        =   158
         Top             =   1440
         Width           =   6975
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   3360
            Picture         =   "MainForm.frx":701A2
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   165
            ToolTipText     =   "Holz"
            Top             =   720
            Width           =   375
            Begin VB.Label Label1 
               Caption         =   "Label1"
               Height          =   375
               Index           =   1
               Left            =   360
               TabIndex        =   166
               Top             =   0
               Width           =   15
            End
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   4440
            Picture         =   "MainForm.frx":705AA
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   164
            ToolTipText     =   "Werkzeug"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   4440
            Picture         =   "MainForm.frx":709C9
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   163
            ToolTipText     =   "Stein"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   5640
            Picture         =   "MainForm.frx":70E13
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   162
            ToolTipText     =   "Glas"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   3360
            Picture         =   "MainForm.frx":7128E
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   161
            ToolTipText     =   "Goldmünzen"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture7 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   5640
            Picture         =   "MainForm.frx":716CD
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   160
            ToolTipText     =   "Mosaik"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox item_pic 
            BorderStyle     =   0  'None
            Height          =   855
            Index           =   1
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   855
            TabIndex        =   159
            Top             =   240
            Width           =   855
         End
         Begin VB.Label out_needed 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   178
            Top             =   360
            Width           =   855
         End
         Begin VB.Label out_geld 
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   177
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_holz 
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   176
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_werkzeug 
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   4920
            TabIndex        =   175
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_stein 
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   4920
            TabIndex        =   174
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_glas 
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   6120
            TabIndex        =   173
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_mosaik 
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   6120
            TabIndex        =   172
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_betrieb 
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   171
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Betriebskosten:"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   170
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Benötigt:"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   169
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label out_last 
            Caption         =   "0%"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   168
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label22 
            Caption         =   "Auslastung:"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   167
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame out_Frame 
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   6975
         Begin VB.PictureBox item_pic 
            BorderStyle     =   0  'None
            Height          =   855
            Index           =   0
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   855
            TabIndex        =   72
            Top             =   240
            Width           =   855
         End
         Begin VB.PictureBox Picture7 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   5640
            Picture         =   "MainForm.frx":71B4F
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   56
            ToolTipText     =   "Mosaik"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   3360
            Picture         =   "MainForm.frx":71FD1
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   55
            ToolTipText     =   "Goldmünzen"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   5640
            Picture         =   "MainForm.frx":72410
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   54
            ToolTipText     =   "Glas"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   4440
            Picture         =   "MainForm.frx":7288B
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   53
            ToolTipText     =   "Stein"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   4440
            Picture         =   "MainForm.frx":72CD5
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   52
            ToolTipText     =   "Werkzeug"
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   3360
            Picture         =   "MainForm.frx":730F4
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   50
            ToolTipText     =   "Holz"
            Top             =   720
            Width           =   375
            Begin VB.Label Label1 
               Caption         =   "Label1"
               Height          =   375
               Index           =   0
               Left            =   360
               TabIndex        =   51
               Top             =   0
               Width           =   15
            End
         End
         Begin VB.Label Label22 
            Caption         =   "Auslastung:"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   157
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label out_last 
            Caption         =   "0%"
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   156
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Benötigt:"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   66
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Betriebskosten:"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   65
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label out_betrieb 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   64
            Top             =   840
            Width           =   855
         End
         Begin VB.Label out_mosaik 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   6120
            TabIndex        =   63
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_glas 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   6120
            TabIndex        =   62
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_stein 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   4920
            TabIndex        =   61
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_werkzeug 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   4920
            TabIndex        =   60
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_holz 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   59
            Top             =   720
            Width           =   735
         End
         Begin VB.Label out_geld 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   58
            Top             =   360
            Width           =   735
         End
         Begin VB.Label out_needed 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   57
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Label error_item 
         Caption         =   "Du musst ein Produkt auswählen für das Die Produktionskette berechnet werden soll."
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
         Height          =   615
         Left            =   240
         TabIndex        =   136
         Top             =   360
         Width           =   6615
      End
      Begin VB.Label welcome_description 
         Caption         =   $"MainForm.frx":734FC
         Height          =   1455
         Left            =   240
         TabIndex        =   135
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Frame OxidentFrame 
      Caption         =   "Waren für den Okzident"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3855
      Begin VB.OptionButton chk_Kurschner 
         Caption         =   "Pelzmäntel"
         Height          =   255
         Left            =   2160
         TabIndex        =   286
         Top             =   4080
         Width           =   1215
      End
      Begin VB.OptionButton chk_Fisch 
         Caption         =   "Fisch"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton chk_Most 
         Caption         =   "Most"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   735
      End
      Begin VB.OptionButton chk_Gewurze 
         Caption         =   "Gewürze"
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton chk_Leinenkutten 
         Caption         =   "Leinenkutten"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
      End
      Begin VB.OptionButton chk_Backhaus 
         Caption         =   "Brot"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton chk_Klosterbrauerei 
         Caption         =   "Bier"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1095
      End
      Begin VB.OptionButton chk_Gerberei 
         Caption         =   "Lederwämser"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3360
         Width           =   1335
      End
      Begin VB.OptionButton chk_Druckerei 
         Caption         =   "Bücher"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3720
         Width           =   1575
      End
      Begin VB.OptionButton chk_Feinschmiede 
         Caption         =   "Kerzenhalter "
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   4080
         Width           =   1215
      End
      Begin VB.OptionButton chk_Schlachterei 
         Caption         =   "Fleisch"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   2640
         Width           =   855
      End
      Begin VB.OptionButton chk_Kelterhaus 
         Caption         =   "Wein"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   3000
         Width           =   975
      End
      Begin VB.OptionButton chk_Seidenweberei 
         Caption         =   "Brokatgewänder "
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   3360
         Width           =   1575
      End
      Begin VB.OptionButton chk_Brillenmacherei 
         Caption         =   "Brillen"
         Height          =   195
         Left            =   2160
         TabIndex        =   18
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox anz_adlige 
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox anz_patrizier 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox anz_burger 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox anz_bauern 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox anz_bettler 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Bauern / Bettler"
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
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Bürger"
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
         Left            =   2040
         TabIndex        =   39
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Patrizier"
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
         Left            =   120
         TabIndex        =   38
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Adlige"
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
         Left            =   2040
         TabIndex        =   37
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Adlige"
         Height          =   255
         Left            =   3000
         TabIndex        =   35
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Patrizier"
         Height          =   255
         Left            =   2280
         TabIndex        =   34
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Bürger"
         Height          =   255
         Left            =   1560
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Bauern"
         Height          =   255
         Left            =   840
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Bettler"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame OrientFrame 
      Caption         =   "Waren für den Orient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   41
      Top             =   1680
      Width           =   3855
      Begin VB.OptionButton chk_Duftmischerei 
         Caption         =   "Duftwasser"
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   2520
         Width           =   1455
      End
      Begin VB.OptionButton chk_Perlenknupferei 
         Caption         =   "Perlenketten"
         Height          =   255
         Left            =   2160
         TabIndex        =   26
         Top             =   2160
         Width           =   1455
      End
      Begin VB.OptionButton chk_Rosterei 
         Caption         =   "Kaffee"
         Height          =   255
         Left            =   2160
         TabIndex        =   25
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton chk_Zuckerbackerei 
         Caption         =   "Marzipan"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton chk_Teppichknupferei 
         Caption         =   "Teppiche"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton chk_Ziegenfarm 
         Caption         =   "Milch"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton chk_Dattelplantage 
         Caption         =   "Datteln"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox anz_gesandte 
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox anz_nomaden 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Gesandte"
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
         Left            =   2040
         TabIndex        =   45
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Nomaden"
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
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Gesandte"
         Height          =   255
         Left            =   2040
         TabIndex        =   43
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Nomaden"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame SonstigeFrame 
      Caption         =   "Sonstige"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   137
      Top             =   1680
      Width           =   3855
      Begin VB.TextBox anz_sonstige 
         Height          =   285
         Left            =   120
         TabIndex        =   151
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton chk_Kanonengiesserei 
         Caption         =   "Kanonen"
         Height          =   255
         Left            =   2160
         TabIndex        =   143
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton chk_Kriegsmaschinenwerkstatt 
         Caption         =   "Kriegsmaschinen"
         Height          =   255
         Left            =   2160
         TabIndex        =   142
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton chk_Glasschmelze 
         Caption         =   "Glas"
         Height          =   255
         Left            =   240
         TabIndex        =   141
         Top             =   2160
         Width           =   855
      End
      Begin VB.OptionButton chk_Waffenschmiede 
         Caption         =   "Waffen"
         Height          =   255
         Left            =   2160
         TabIndex        =   140
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton chk_Seilerei 
         Caption         =   "Seile"
         Height          =   255
         Left            =   240
         TabIndex        =   139
         Top             =   1800
         Width           =   975
      End
      Begin VB.OptionButton chk_Werkzeugmacher 
         Caption         =   "Werkzeug"
         Height          =   255
         Left            =   240
         TabIndex        =   138
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Kriegswerkzeug"
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
         Left            =   2040
         TabIndex        =   155
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "Baumaterial"
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
         Left            =   120
         TabIndex        =   154
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Gebäudeanzahl"
         Height          =   255
         Left            =   120
         TabIndex        =   153
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Menu menuProgramm 
      Caption         =   "Programm"
      Begin VB.Menu menuModus 
         Caption         =   "Modus"
         WindowList      =   -1  'True
         Begin VB.Menu menuOxident 
            Caption         =   "Oxident"
            Checked         =   -1  'True
         End
         Begin VB.Menu menuOrient 
            Caption         =   "Orient"
         End
         Begin VB.Menu menuSonstige 
            Caption         =   "Sonstige"
         End
      End
      Begin VB.Menu MenuEnd 
         Caption         =   "Beenden"
      End
   End
   Begin VB.Menu MenuHilfe 
      Caption         =   "Hilfe"
      Begin VB.Menu MenuHilfe1 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mode
Dim item_name(64)
Dim braucht(64)
Dim versorgt(64)
Dim gesamt(16)
Dim need_sub(64, 9)
Dim Position
Dim i_Fischerhutte
Dim i_Gewurzplantage
Dim i_Mosthof
Dim i_Hanfplantage
Dim i_Leinenkutten
Dim i_Backhaus
Dim i_Muhle
Dim i_Weizenfarm
Dim i_Krauterfarm
Dim i_Klosterbrauerei
Dim i_Gerberei
Dim i_Schweinezucht
Dim i_Saline
Dim i_Salzmine
Dim i_Kohlerhutte
Dim i_Kupfermine
Dim i_Kupferschmelze
Dim i_Holzfallerhutte
Dim i_Papiermuhle
Dim i_Indigoplantage
Dim i_Druckerei
Dim i_Feinschmiede
Dim i_Imkerei
Dim i_Lichtzieherei
Dim i_Rinderfarmen
Dim i_Schlachterei
Dim i_Eisenmine
Dim i_Eisenschmelze
Dim i_Fasskuferei
Dim i_Kelterhaus
Dim i_Weingut
Dim i_Seidenplantage
Dim i_Goldmine
Dim i_Goldschmelze
Dim i_Seidenweberei
Dim i_Quarzbruch
Dim i_Brillenmacherei
Dim i_Teppichknupferei
Dim i_Dattelplantage
Dim i_Ziegenfarm
Dim i_Zuckerrohrplantage
Dim i_Zuckermuhle
Dim i_Zuckerbackerei
Dim i_Mandelplantage
Dim i_Rosenzuchterei
Dim i_Duftmischerei
Dim i_Perlentaucherhutte
Dim i_Perlenknupferei
Dim i_Kaffeeplantage
Dim i_Rosterei
Dim i_Werkzeugmacher
Dim i_Seilerei
Dim i_Glasschmelze
Dim i_Waldglashutte
Dim i_Kriegsmaschinenwerkstatt
Dim i_Kanonengiesserei
Dim i_Waffenschmiede
Dim i_Kurschner
Dim i_Pelztierjager

Private Sub Form_Load()
    init = init_Script()
       
    'braucht(ID) = Array(geld, holz, stein, werkzeug, glas, mosaik, betriebskosten)
    'versorgt(ID) = Array(Bettler, Bauern, Bürger, Patrizier, Adlige, Nomaden, Gesandte, Sonstige)
    'need_sub(ID, count 0-*) = Array(SUB_ID, Anzahl)
    
    item_name(i_Fischerhutte) = "Fischerhütte"
    braucht(i_Fischerhutte) = Array(100, 3, 0, 2, 0, 0, 15)
    versorgt(i_Fischerhutte) = Array(286, 200, 500, 909, 1250, 0, 0)
    
    item_name(i_Mosthof) = "Mosthof"
    braucht(i_Mosthof) = Array(100, 5, 0, 1, 0, 0, 15)
    versorgt(i_Mosthof) = Array(500, 341, 341, 625, 1154, 0, 0)
    
    item_name(i_Gewurzplantage) = "Gewürzplantage"
    braucht(i_Gewurzplantage) = Array(500, 5, 0, 2, 0, 0, 30)
    versorgt(i_Gewurzplantage) = Array(0, 0, 500, 909, 1250, 0, 0)
    
    item_name(i_Leinenkutten) = "Leinenkutten"
    braucht(i_Leinenkutten) = Array(400, 5, 0, 3, 0, 0, 25)
    versorgt(i_Leinenkutten) = Array(0, 0, 476, 1053, 2500, 0, 0)
    need_sub(i_Leinenkutten, 0) = Array(i_Hanfplantage, 2)

    item_name(i_Hanfplantage) = "Hanfplantage"
    braucht(i_Hanfplantage) = Array(200, 5, 0, 2, 0, 0, 20)

    item_name(i_Backhaus) = "Backhaus"
    braucht(i_Backhaus) = Array(700, 5, 5, 5, 0, 0, 30)
    versorgt(i_Backhaus) = Array(0, 0, 0, 727, 1026, 0, 0)
    need_sub(i_Backhaus, 0) = Array(i_Muhle, 1)
    
    item_name(i_Muhle) = "Mühle"
    braucht(i_Muhle) = Array(800, 8, 4, 4, 0, 0, 30)
    need_sub(i_Muhle, 0) = Array(i_Weizenfarm, 2)

    item_name(i_Weizenfarm) = "Weizenfarm"
    braucht(i_Weizenfarm) = Array(200, 8, 0, 2, 0, 0, 5)

    item_name(i_Klosterbrauerei) = "Klosterbrauerei"
    braucht(i_Klosterbrauerei) = Array(600, 6, 4, 5, 0, 0, 30)
    versorgt(i_Klosterbrauerei) = Array(0, 0, 0, 625, 1071, 0, 0)
    need_sub(i_Klosterbrauerei, 0) = Array(i_Weizenfarm, 1)
    need_sub(i_Klosterbrauerei, 1) = Array(i_Krauterfarm, 1)

    item_name(i_Krauterfarm) = "Kräuterfarm"
    braucht(i_Krauterfarm) = Array(200, 5, 4, 2, 0, 0, 10)

    item_name(i_Gerberei) = "Gerberei"
    braucht(i_Gerberei) = Array(700, 7, 8, 3, 0, 0, 20)
    versorgt(i_Gerberei) = Array(0, 0, 0, 1429, 2500, 0, 0)
    need_sub(i_Gerberei, 0) = Array(i_Schweinezucht, 2)
    need_sub(i_Gerberei, 1) = Array(i_Saline, 1)

    item_name(i_Schweinezucht) = "Schweinezucht"
    braucht(i_Schweinezucht) = Array(400, 7, 8, 3, 0, 0, 15)

    item_name(i_Saline) = "Saline"
    braucht(i_Saline) = Array(900, 3, 6, 5, 0, 0, 25)
    need_sub(i_Saline, 0) = Array(i_Salzmine, 1)
    need_sub(i_Saline, 1) = Array(i_Kohlerhutte, 1)

    item_name(i_Salzmine) = "Salzmine"
    braucht(i_Salzmine) = Array(800, 11, 5, 4, 0, 0, 20)

    item_name(i_Kohlerhutte) = "Köhlerhütte"
    braucht(i_Kohlerhutte) = Array(250, 3, 2, 2, 0, 0, 10)

    item_name(i_Druckerei) = "Druckerei"
    braucht(i_Druckerei) = Array(1800, 5, 12, 5, 10, 0, 50)
    versorgt(i_Druckerei) = Array(0, 0, 0, 1875, 3333, 0, 0)
    need_sub(i_Druckerei, 0) = Array(i_Indigoplantage, 2)
    need_sub(i_Druckerei, 1) = Array(i_Papiermuhle, 0.5)

    item_name(i_Indigoplantage) = "Indigoplantage"
    braucht(i_Indigoplantage) = Array(400, 5, 0, 2, 0, 0, 20)

    item_name(i_Papiermuhle) = "Papiermühle"
    braucht(i_Papiermuhle) = Array(1500, 5, 12, 5, 0, 0, 50)
    need_sub(i_Papiermuhle, 0) = Array(i_Holzfallerhutte, 2)

    item_name(i_Holzfallerhutte) = "Holzfällerhütte"
    braucht(i_Holzfallerhutte) = Array(50, 0, 0, 1, 0, 0, 20)

    item_name(i_Feinschmiede) = "Feinschmiede"
    braucht(i_Feinschmiede) = Array(1800, 9, 15, 7, 10, 0, 60)
    versorgt(i_Feinschmiede) = Array(0, 0, 0, 2500, 3333, 0, 0)
    need_sub(i_Feinschmiede, 0) = Array(i_Kupferschmelze, 0.55)
    need_sub(i_Feinschmiede, 1) = Array(i_Lichtzieherei, 1.4)

    item_name(i_Kupferschmelze) = "Kupferschmelze"
    braucht(i_Kupferschmelze) = Array(1500, 5, 12, 5, 0, 0, 50)
    need_sub(i_Kupferschmelze, 0) = Array(i_Kohlerhutte, 0.7)
    need_sub(i_Kupferschmelze, 1) = Array(i_Kupfermine, 1)

    item_name(i_Kupfermine) = "Kupfermine"
    braucht(i_Kupfermine) = Array(1500, 10, 12, 8, 0, 0, 40)

    item_name(i_Lichtzieherei) = "Lichtzieherei"
    braucht(i_Lichtzieherei) = Array(1600, 7, 10, 6, 0, 0, 40)
    need_sub(i_Lichtzieherei, 0) = Array(i_Hanfplantage, 1)
    need_sub(i_Lichtzieherei, 1) = Array(i_Imkerei, 2)

    item_name(i_Imkerei) = "Imkerei"
    braucht(i_Imkerei) = Array(500, 7, 9, 3, 0, 0, 15)
    
    item_name(i_Schlachterei) = "Schlachterei"
    braucht(i_Schlachterei) = Array(1000, 5, 8, 7, 0, 0, 50)
    versorgt(i_Schlachterei) = Array(0, 0, 0, 0, 1136, 0, 0)
    need_sub(i_Schlachterei, 0) = Array(i_Saline, 0.48)
    need_sub(i_Schlachterei, 1) = Array(i_Rinderfarmen, 2)
    
    item_name(i_Rinderfarmen) = "Rinderfarmen"
    braucht(i_Rinderfarmen) = Array(600, 8, 6, 2, 0, 0, 25)
    
    item_name(i_Kelterhaus) = "Kelterhaus"
    braucht(i_Kelterhaus) = Array(1800, 14, 7, 7, 9, 0, 50)
    versorgt(i_Kelterhaus) = Array(0, 0, 0, 0, 1000, 0, 0)
    need_sub(i_Kelterhaus, 0) = Array(i_Weingut, 3)
    need_sub(i_Kelterhaus, 1) = Array(i_Fasskuferei, 1)
    
    item_name(i_Weingut) = "Weingut"
    braucht(i_Weingut) = Array(800, 8, 10, 4, 0, 0, 25)
    
    item_name(i_Fasskuferei) = "Fassküferei"
    braucht(i_Fasskuferei) = Array(1000, 7, 8, 5, 0, 0, 30)
    need_sub(i_Fasskuferei, 0) = Array(i_Holzfallerhutte, 0.66)
    need_sub(i_Fasskuferei, 1) = Array(i_Eisenschmelze, 0.5)
    
    item_name(i_Eisenschmelze) = "Eisenschmelze"
    braucht(i_Eisenschmelze) = Array(600, 10, 2, 5, 0, 0, 20)
    need_sub(i_Eisenschmelze, 0) = Array(i_Kohlerhutte, 1)
    need_sub(i_Eisenschmelze, 1) = Array(i_Eisenmine, 1)
    
    item_name(i_Eisenmine) = "Eisenmine"
    braucht(i_Eisenmine) = Array(900, 12, 2, 5, 0, 0, 20)
    
    item_name(i_Seidenweberei) = "Seidenweberei"
    braucht(i_Seidenweberei) = Array(1500, 5, 12, 8, 15, 0, 80)
    versorgt(i_Seidenweberei) = Array(0, 0, 0, 0, 2112, 0, 0)
    need_sub(i_Seidenweberei, 0) = Array(i_Goldschmelze, 1)
    need_sub(i_Seidenweberei, 1) = Array(i_Seidenplantage, 0.5)

    item_name(i_Goldschmelze) = "Goldschmelze"
    braucht(i_Goldschmelze) = Array(2000, 14, 11, 16, 0, 0, 30)
    need_sub(i_Goldschmelze, 0) = Array(i_Kohlerhutte, 0.75)
    need_sub(i_Goldschmelze, 1) = Array(i_Goldmine, 1)

    item_name(i_Goldmine) = "Goldmine"
    braucht(i_Goldmine) = Array(2500, 20, 12, 13, 0, 0, 50)
    
    item_name(i_Seidenplantage) = "Seidenplantage"
    braucht(i_Seidenplantage) = Array(350, 5, 0, 2, 0, 0, 25)
     
    item_name(i_Brillenmacherei) = "Brillenmacherei"
    braucht(i_Brillenmacherei) = Array(1800, 8, 14, 6, 15, 0, 40)
    versorgt(i_Brillenmacherei) = Array(0, 0, 0, 0, 1709, 0, 0)
    need_sub(i_Brillenmacherei, 0) = Array(i_Kupferschmelze, 0.75)
    need_sub(i_Brillenmacherei, 1) = Array(i_Quarzbruch, 0.75)
    
    item_name(i_Quarzbruch) = "Quarzbruch"
    braucht(i_Quarzbruch) = Array(1000, 10, 0, 6, 0, 0, 20)
  
    item_name(i_Dattelplantage) = "Dattelplantage"
    braucht(i_Dattelplantage) = Array(200, 3, 0, 2, 0, 0, 45)
    versorgt(i_Dattelplantage) = Array(0, 0, 0, 0, 0, 450, 600)
  
    item_name(i_Ziegenfarm) = "Ziegenfarm"
    braucht(i_Ziegenfarm) = Array(200, 5, 0, 1, 0, 0, 20)
    versorgt(i_Ziegenfarm) = Array(0, 0, 0, 0, 0, 436, 667)
    
    item_name(i_Teppichknupferei) = "Teppichknüpferei"
    braucht(i_Teppichknupferei) = Array(400, 5, 0, 3, 0, 0, 60)
    versorgt(i_Teppichknupferei) = Array(0, 0, 0, 0, 0, 909, 1500)
    need_sub(i_Teppichknupferei, 0) = Array(i_Seidenplantage, 1)
    need_sub(i_Teppichknupferei, 1) = Array(i_Indigoplantage, 1)
    
    item_name(i_Zuckerbackerei) = "Zuckerbäckerei"
    braucht(i_Zuckerbackerei) = Array(1500, 5, 0, 12, 0, 24, 100)
    versorgt(i_Zuckerbackerei) = Array(0, 0, 0, 0, 0, 0, 2454)
    need_sub(i_Zuckerbackerei, 0) = Array(i_Zuckermuhle, 0.5)
    need_sub(i_Zuckerbackerei, 1) = Array(i_Mandelplantage, 2)

    item_name(i_Mandelplantage) = "Mandelplantage"
    braucht(i_Mandelplantage) = Array(500, 4, 0, 6, 0, 10, 40)

    item_name(i_Zuckermuhle) = "Zuckermühle"
    braucht(i_Zuckermuhle) = Array(800, 7, 0, 6, 0, 10, 40)
    need_sub(i_Zuckermuhle, 0) = Array(i_Zuckerrohrplantage, 2)

    item_name(i_Zuckerrohrplantage) = "Zuckerrohrplantage"
    braucht(i_Zuckerrohrplantage) = Array(500, 5, 0, 3, 0, 9, 35)
    
    item_name(i_Rosterei) = "Rösterei"
    braucht(i_Rosterei) = Array(1100, 5, 0, 10, 0, 15, 45)
    versorgt(i_Rosterei) = Array(0, 0, 0, 0, 0, 0, 1000)
    need_sub(i_Rosterei, 0) = Array(i_Kaffeeplantage, 2)

    item_name(i_Kaffeeplantage) = "Kaffeeplantage"
    braucht(i_Kaffeeplantage) = Array(500, 2, 0, 4, 0, 6, 20)
    
    item_name(i_Perlenknupferei) = "Perlenknüpferei"
    braucht(i_Perlenknupferei) = Array(1800, 8, 0, 8, 0, 16, 70)
    versorgt(i_Perlenknupferei) = Array(0, 0, 0, 0, 0, 0, 752)
    need_sub(i_Perlenknupferei, 0) = Array(i_Perlentaucherhutte, 1)

    item_name(i_Perlentaucherhutte) = "Perlentaucherhütte"
    braucht(i_Perlentaucherhutte) = Array(1200, 14, 0, 7, 0, 11, 40)
    
    item_name(i_Duftmischerei) = "Duftmischerei"
    braucht(i_Duftmischerei) = Array(2500, 12, 0, 9, 0, 24, 60)
    versorgt(i_Duftmischerei) = Array(0, 0, 0, 0, 0, 0, 1250)
    need_sub(i_Duftmischerei, 0) = Array(i_Rosenzuchterei, 3)

    item_name(i_Rosenzuchterei) = "Rosenzüchterei"
    braucht(i_Rosenzuchterei) = Array(900, 10, 0, 5, 0, 12, 30)
    
    item_name(i_Werkzeugmacher) = "Werkzeugmacher"
    braucht(i_Werkzeugmacher) = Array(500, 8, 2, 5, 0, 0, 30)
    need_sub(i_Werkzeugmacher, 0) = Array(i_Eisenschmelze, 0.5)
    
    item_name(i_Seilerei) = "Seilerei"
    braucht(i_Seilerei) = Array(700, 12, 5, 0, 0, 0, 40)
    need_sub(i_Seilerei, 0) = Array(i_Hanfplantage, 1)

    item_name(i_Glasschmelze) = "Glasschmelze"
    braucht(i_Glasschmelze) = Array(1200, 10, 12, 5, 0, 0, 30)
    need_sub(i_Glasschmelze, 0) = Array(i_Quarzbruch, 0.75)
    need_sub(i_Glasschmelze, 1) = Array(i_Waldglashutte, 1)
    
    item_name(i_Waldglashutte) = "Waldglashütte"
    braucht(i_Waldglashutte) = Array(500, 6, 8, 4, 0, 0, 30)

    item_name(i_Kriegsmaschinenwerkstatt) = "Kriegsmaschinenwerkstatt"
    braucht(i_Kriegsmaschinenwerkstatt) = Array(3000, 3, 10, 5, 8, 0, 60)
    need_sub(i_Kriegsmaschinenwerkstatt, 0) = Array(i_Holzfallerhutte, 0.5)
    need_sub(i_Kriegsmaschinenwerkstatt, 1) = Array(i_Seilerei, 0.75)

    item_name(i_Kanonengiesserei) = "Kanonengießerei"
    braucht(i_Kanonengiesserei) = Array(6000, 24, 30, 15, 24, 0, 100)
    need_sub(i_Kanonengiesserei, 0) = Array(i_Holzfallerhutte, 1)
    need_sub(i_Kanonengiesserei, 1) = Array(i_Eisenschmelze, 1)

    item_name(i_Waffenschmiede) = "Waffenschmiede"
    braucht(i_Waffenschmiede) = Array(1500, 3, 10, 5, 24, 0, 30)
    need_sub(i_Waffenschmiede, 0) = Array(i_Eisenschmelze, 1)
       
       
    item_name(i_Kurschner) = "Kurschner"
    braucht(i_Kurschner) = Array(1600, 3, 10, 8, 0, 0, 90)
    versorgt(i_Kurschner) = Array(0, 0, 0, 0, 1563, 0, 0)
    need_sub(i_Kurschner, 0) = Array(i_Saline, 0.33)
    need_sub(i_Kurschner, 1) = Array(i_Pelztierjager, 1)
  
    item_name(i_Pelztierjager) = "Pelztierjäger"
    braucht(i_Pelztierjager) = Array(900, 7, 4, 2, 0, 0, 30)

End Sub

Public Function init_Script()
    i_Fischerhutte = 1
    i_Hanfplantage = 2
    i_Mosthof = 3
    i_Gewurzplantage = 4
    i_Leinenkutten = 5
    i_Backhaus = 6
    i_Muhle = 7
    i_Weizenfarm = 8
    i_Klosterbrauerei = 9
    i_Krauterfarm = 10
    i_Gerberei = 11
    i_Schweinezucht = 12
    i_Saline = 13
    i_Salzmine = 14
    i_Kohlerhutte = 15
    i_Kupfermine = 16
    i_Kupferschmelze = 17
    i_Holzfallerhutte = 18
    i_Papiermuhle = 19
    i_Indigoplantage = 20
    i_Druckerei = 21
    i_Feinschmiede = 22
    i_Imkerei = 23
    i_Lichtzieherei = 24
    i_Rinderfarmen = 25
    i_Schlachterei = 26
    i_Eisenmine = 27
    i_Eisenschmelze = 28
    i_Fasskuferei = 29
    i_Kelterhaus = 30
    i_Weingut = 31
    i_Seidenplantage = 32
    i_Goldmine = 33
    i_Goldschmelze = 34
    i_Seidenweberei = 35
    i_Brillenmacherei = 36
    i_Quarzbruch = 37
    i_Teppichknupferei = 38
    i_Dattelplantage = 39
    i_Ziegenfarm = 40
    i_Zuckerrohrplantage = 41
    i_Zuckermuhle = 42
    i_Zuckerbackerei = 43
    i_Mandelplantage = 44
    i_Rosenzuchterei = 45
    i_Duftmischerei = 46
    i_Perlentaucherhutte = 47
    i_Perlenknupferei = 48
    i_Kaffeeplantage = 49
    i_Rosterei = 50
    i_Werkzeugmacher = 51
    i_Seilerei = 52
    i_Glasschmelze = 53
    i_Waldglashutte = 54
    i_Kriegsmaschinenwerkstatt = 55
    i_Kanonengiesserei = 56
    i_Waffenschmiede = 57
    i_Kurschner = 58
    i_Pelztierjager = 59


    Position = 1
    Mode = 0

    gesamt(0) = 0
    gesamt(1) = 0
    gesamt(2) = 0
    gesamt(3) = 0
    gesamt(4) = 0
    gesamt(5) = 0
    gesamt(6) = 0
    
    outFrame.Caption = "Daten eingeben"
    error_item.Visible = False
    Me.KeyPreview = True
    
    out_Frame(0).Visible = False
    out_Frame(1).Visible = False
    out_Frame(2).Visible = False
    out_Frame(3).Visible = False
    out_Frame(4).Visible = False
    out_Frame(5).Visible = False
    out_Frame(6).Visible = False
End Function

Public Function Output_needed(Position, item, needed)
    out_Frame(Position).Visible = True
    out_Frame(Position).Caption = item_name(item)
    item_pic(Position).Picture = pic(item).Picture

    out_needed(Position).Caption = needed
    gesamt(0) = gesamt(0) + needed * braucht(item)(0)
    gesamt(1) = gesamt(1) + needed * braucht(item)(1)
    gesamt(2) = gesamt(2) + needed * braucht(item)(2)
    gesamt(3) = gesamt(3) + needed * braucht(item)(3)
    gesamt(4) = gesamt(4) + needed * braucht(item)(4)
    gesamt(5) = gesamt(5) + needed * braucht(item)(5)
    gesamt(6) = gesamt(6) + needed * braucht(item)(6)
    
    out_geld(Position).Caption = needed * braucht(item)(0)
    out_holz(Position).Caption = needed * braucht(item)(1)
    out_stein(Position).Caption = needed * braucht(item)(2)
    out_werkzeug(Position).Caption = needed * braucht(item)(3)
    out_glas(Position).Caption = needed * braucht(item)(4)
    out_mosaik(Position).Caption = needed * braucht(item)(5)
    out_betrieb(Position).Caption = needed * braucht(item)(6)
End Function

Public Function Output_sub(item, needed, top_needed)
    
    If Round(needed) < needed Then
        new1_needed = needed + 0.5
    Else
        new1_needed = needed
    End If
    new_needed = Round(new1_needed)
    If new_needed < 1 Then
        new_needed = 1
    End If
    
    out = Output_needed(Position, item(0), new_needed)
    leistung = (needed / new_needed) * 100
    leistung = Round(leistung) & "%"
    out_last(Position).Caption = leistung

    Position = Position + 1
    If IsArray(need_sub(item(0), 0)) Then
        i = 0
        While IsArray(need_sub(item(0), i))

            out_sub = Output_sub(need_sub(item(0), i), (need_sub(item(0), i)(1) * new_needed), new_needed)
             i = i + 1
        Wend
    End If
End Function

Public Function DoOutput(item)
    Dim needed As Double
    'needed = 0
    outFrame.Caption = "Produktionskette"

    If Not IsNumeric(anz_bettler.Text) Then
          anz_bettler.Text = 0
    End If
    If Not IsNumeric(anz_bauern.Text) Then
          anz_bauern.Text = 0
    End If
    If Not IsNumeric(anz_burger.Text) Then
          anz_burger.Text = 0
    End If
    If Not IsNumeric(anz_patrizier.Text) Then
          anz_patrizier.Text = 0
    End If
    If Not IsNumeric(anz_adlige.Text) Then
          anz_adlige.Text = 0
    End If
    If Not IsNumeric(anz_nomaden.Text) Then
          anz_nomaden.Text = 0
    End If
    If Not IsNumeric(anz_gesandte.Text) Then
          anz_gesandte.Text = 0
    End If
    If Not IsNumeric(anz_sonstige.Text) Then
          anz_sonstige.Text = 0
    End If
    
    If Mode = 1 Then
        needed = anz_sonstige.Text
        'Debug.Print needed
    Else
        If versorgt(item)(0) > 0 Then
            needed = (anz_bettler.Text / versorgt(item)(0))
        End If
        If versorgt(item)(1) > 0 Then
            needed = needed + (anz_bauern.Text / versorgt(item)(1))
        End If
        If versorgt(item)(2) > 0 Then
            needed = needed + (anz_burger.Text / versorgt(item)(2))
        End If
        If versorgt(item)(3) > 0 Then
            needed = needed + (anz_patrizier.Text / versorgt(item)(3))
        End If
        If versorgt(item)(4) > 0 Then
            needed = needed + (anz_adlige.Text / versorgt(item)(4))
        End If
        If versorgt(item)(5) > 0 Then
            needed = needed + (anz_nomaden.Text / versorgt(item)(5))
        End If
        If versorgt(item)(6) > 0 Then
            needed = needed + (anz_gesandte.Text / versorgt(item)(6))
        End If
    End If
    
    If needed < 1 And needed > 0.00000001 Then
        needed = 0.6
    End If
    'needed = Round(needed)
    
    
    ganz_needed = Round(needed)
    'Debug.Print needed
    If (ganz_needed < needed) Then
        ganz_needed = ganz_needed + 1
    Else
        ganz_needed = ganz_needed
    End If
    'needed = Round(needed + 0.5)
    
    
    'a = (versorgt(item)(0) + versorgt(item)(1) + versorgt(item)(2) + versorgt(item)(3) + versorgt(item)(4) + versorgt(item)(5) + versorgt(item)(6))
    'b = anz_bettler.Text + anz_burger.Text + anz_patrizier.Text + anz_adlige.Text + anz_nomaden.Text + anz_gesandte.Text
    
    If IsArray(versorgt(item)) Then
    
    
        If versorgt(item)(0) > 0 Then
            v_bettler = (anz_bettler.Text / versorgt(item)(0))
        Else
            v_bettler = 0
        End If
        If versorgt(item)(1) > 0 Then
            v_bauern = (anz_bauern.Text / versorgt(item)(1))
        Else
            v_bauern = 0
        End If
        If versorgt(item)(2) > 0 Then
            v_burger = (anz_burger.Text / versorgt(item)(2))
        Else
            v_burger = 0
        End If
        If versorgt(item)(3) > 0 Then
            v_patrizier = (anz_patrizier.Text / versorgt(item)(3))
        Else
            v_patrizier = 0
        End If
        If versorgt(item)(4) > 0 Then
            v_adlige = (anz_adlige.Text / versorgt(item)(4))
        Else
            v_adlige = 0
        End If
        If versorgt(item)(5) > 0 Then
            v_nomaden = (anz_nomaden.Text / versorgt(item)(5))
        Else
            v_nomaden = 0
        End If
        If versorgt(item)(6) > 0 Then
            v_gesandte = (anz_gesandte.Text / versorgt(item)(6))
        Else
            v_gesandte = 0
        End If
        
        If needed > 0 Then
            v_gesamt = (v_bettler + v_bauern + v_burger + v_patrizier + v_adlige + v_nomaden + v_gesandte) / ganz_needed
        End If
    Else
        v_gesamt = 1
    End If

    v_gesamt = Round(v_gesamt * 100)
    
    o = v_gesamt & "%"
    
    out_last(0).Caption = o

    out = Output_needed(0, item, ganz_needed)
    
    If IsArray(need_sub(item, 0)) And needed > 0 Then
        i = 0
        While IsArray(need_sub(item, i))
            new_needed = (need_sub(item, i)(1) * ganz_needed)
            
                If Round(new_needed) < new_needed Then
                    new1_needed = new_needed + 0.5
                Else
                    new1_needed = new_needed
                End If
            
            top_needed = (new1_needed / Round(new1_needed)) * 100
            'out_sub = Output_sub(need_sub(item, i), new_needed, need_sub(item, i)(1))
            out_sub = Output_sub(need_sub(item, i), new_needed, ganz_needed)
            i = i + 1
        Wend
    End If
End Function

Public Function show_Sonstige()
    menuOrient.Checked = False
    menuOxident.Checked = False
    menuSonstige.Checked = True
    SonstigeFrame.Visible = True
    OxidentFrame.Visible = False
    OrientFrame.Visible = False
    chk_Rosterei.Value = False
    chk_Perlenknupferei.Value = False
    chk_Duftmischerei.Value = False
    chk_Zuckerbackerei.Value = False
    chk_Ziegenfarm.Value = False
    chk_Dattelplantage.Value = False
    chk_Teppichknupferei.Value = False
    chk_Fisch.Value = False
    chk_Most.Value = False
    chk_Gewurze = False
    chk_Leinenkutten = False
    chk_Backhaus = False
    chk_Klosterbrauerei = False
    chk_Gerberei = False
    chk_Druckerei = False
    chk_Feinschmiede = False
    chk_Schlachterei = False
    chk_Kelterhaus = False
    chk_Kurschner = False
    chk_Seidenweberei = False
    chk_Brillenmacherei = False
End Function

Public Function show_Okzident()
    menuOrient.Checked = False
    menuOxident.Checked = True
    menuSonstige.Checked = False
    SonstigeFrame.Visible = False
    OxidentFrame.Visible = True
    OrientFrame.Visible = False
    chk_Rosterei.Value = False
    chk_Perlenknupferei.Value = False
    chk_Duftmischerei.Value = False
    chk_Zuckerbackerei.Value = False
    chk_Ziegenfarm.Value = False
    chk_Dattelplantage.Value = False
    chk_Teppichknupferei.Value = False
    chk_Seilerei = False
    chk_Werkzeugmacher = False
    chk_Waffenschmiede = False
    chk_Kriegsmaschinenwerkstatt = False
    chk_Glasschmelze = False
    chk_Kanonengiesserei = False
End Function

Public Function show_Orient()
    menuOrient.Checked = True
    menuOxident.Checked = False
    menuSonstige.Checked = False
    OxidentFrame.Visible = False
    SonstigeFrame.Visible = False
    OrientFrame.Visible = True
    chk_Fisch.Value = False
    chk_Most.Value = False
    chk_Gewurze = False
    chk_Leinenkutten = False
    chk_Backhaus = False
    chk_Klosterbrauerei = False
    chk_Gerberei = False
    chk_Druckerei = False
    chk_Feinschmiede = False
    chk_Schlachterei = False
    chk_Kelterhaus = False
    chk_Kurschner = False
    chk_Seidenweberei = False
    chk_Brillenmacherei = False
    chk_Seilerei = False
    chk_Werkzeugmacher = False
    chk_Waffenschmiede = False
    chk_Kriegsmaschinenwerkstatt = False
    chk_Glasschmelze = False
    chk_Kanonengiesserei = False
End Function

Function Start_Action()
    init = init_Script()

    If chk_Fisch.Value = True Then
        Output = DoOutput(i_Fischerhutte)
    ElseIf chk_Most.Value = True Then
        Output = DoOutput(i_Mosthof)
    ElseIf chk_Gewurze.Value = True Then
        Output = DoOutput(i_Gewurzplantage)
    ElseIf chk_Leinenkutten.Value = True Then
        Output = DoOutput(i_Leinenkutten)
    ElseIf chk_Backhaus.Value = True Then
        Output = DoOutput(i_Backhaus)
    ElseIf chk_Klosterbrauerei.Value = True Then
        Output = DoOutput(i_Klosterbrauerei)
    ElseIf chk_Gerberei.Value = True Then
        Output = DoOutput(i_Gerberei)
    ElseIf chk_Druckerei.Value = True Then
        Output = DoOutput(i_Druckerei)
    ElseIf chk_Feinschmiede.Value = True Then
        Output = DoOutput(i_Feinschmiede)
    ElseIf chk_Schlachterei.Value = True Then
        Output = DoOutput(i_Schlachterei)
    ElseIf chk_Kelterhaus.Value = True Then
        Output = DoOutput(i_Kelterhaus)
    ElseIf chk_Seidenweberei.Value = True Then
        Output = DoOutput(i_Seidenweberei)
    ElseIf chk_Brillenmacherei.Value = True Then
        Output = DoOutput(i_Brillenmacherei)
    ElseIf chk_Kurschner.Value = True Then
        Output = DoOutput(i_Kurschner)
    ElseIf chk_Teppichknupferei.Value = True Then
        Output = DoOutput(i_Teppichknupferei)
    ElseIf chk_Dattelplantage.Value = True Then
        Output = DoOutput(i_Dattelplantage)
    ElseIf chk_Ziegenfarm.Value = True Then
        Output = DoOutput(i_Ziegenfarm)
    ElseIf chk_Zuckerbackerei.Value = True Then
        Output = DoOutput(i_Zuckerbackerei)
    ElseIf chk_Duftmischerei.Value = True Then
        Output = DoOutput(i_Duftmischerei)
    ElseIf chk_Perlenknupferei.Value = True Then
        Output = DoOutput(i_Perlenknupferei)
    ElseIf chk_Rosterei.Value = True Then
        Output = DoOutput(i_Rosterei)
    ElseIf chk_Seilerei.Value = True Then
        Mode = 1
        Output = DoOutput(i_Seilerei)
    ElseIf chk_Werkzeugmacher.Value = True Then
        Mode = 1
        Output = DoOutput(i_Werkzeugmacher)
    ElseIf chk_Waffenschmiede.Value = True Then
        Mode = 1
        Output = DoOutput(i_Waffenschmiede)
    ElseIf chk_Kriegsmaschinenwerkstatt.Value = True Then
        Mode = 1
        Output = DoOutput(i_Kriegsmaschinenwerkstatt)
    ElseIf chk_Glasschmelze.Value = True Then
        Mode = 1
        Output = DoOutput(i_Glasschmelze)
    ElseIf chk_Kanonengiesserei.Value = True Then
        Mode = 1
        Output = DoOutput(i_Kanonengiesserei)
    Else
        outFrame.Caption = "Fehler"
        error_item.Visible = True
    End If
    
    welcome_description.Visible = False
    
    gesamt_geld.Caption = gesamt(0)
    gesamt_holz.Caption = gesamt(1)
    gesamt_stein.Caption = gesamt(2)
    gesamt_werkzeug.Caption = gesamt(3)
    gesamt_glas.Caption = gesamt(4)
    gesamt_mosaik.Caption = gesamt(5)
    gesamt_kosten.Caption = gesamt(6)
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    start = Start_Action()
End Sub
Private Sub Command1_Click()
    start = Start_Action()
End Sub
Private Sub MenuEnd_Click()
    End
End Sub
Private Sub MenuHilfe1_Click()
    frmAbout.Show
End Sub
Private Sub menuOxident_Click()
    show_mode = show_Okzident()
End Sub
Private Sub Command2_Click()
    show_mode = show_Okzident()
End Sub
Private Sub Command4_Click()
    show_mode = show_Sonstige()
End Sub
Private Sub menuOrient_Click()
    show_mode = show_Orient()
End Sub
Private Sub Command3_Click()
    show_mode = show_Orient()
End Sub
Private Sub menuSonstige_Click()
    show_mode = show_Sonstige()
End Sub
