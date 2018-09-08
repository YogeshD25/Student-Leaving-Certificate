VERSION 5.00
Begin VB.Form transfer_cer 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Transfer Certificate"
   ClientHeight    =   10800
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   19200
   Icon            =   "transfer_cer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000A&
      Height          =   1095
      Left            =   2760
      TabIndex        =   59
      Top             =   8520
      Width           =   13335
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000E&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5040
         Picture         =   "transfer_cer.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Print the TC"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H8000000E&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6360
         Picture         =   "transfer_cer.frx":190C
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Save the Data"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H8000000E&
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7560
         Picture         =   "transfer_cer.frx":254E
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "close this window"
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8535
      Left            =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      Begin VB.Frame Frame6 
         Height          =   2055
         Left            =   120
         TabIndex        =   66
         Top             =   6360
         Visible         =   0   'False
         Width           =   13095
         Begin VB.ComboBox Combo4 
            BackColor       =   &H0080FF80&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "transfer_cer.frx":3190
            Left            =   10320
            List            =   "transfer_cer.frx":319A
            TabIndex        =   80
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox Text25 
            BackColor       =   &H00FF8080&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   2160
            TabIndex        =   78
            Text            =   "AT HIS OWN REQUEST"
            Top             =   1560
            Width           =   2895
         End
         Begin VB.ComboBox Combo3 
            BackColor       =   &H0080FF80&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "transfer_cer.frx":31BF
            Left            =   2160
            List            =   "transfer_cer.frx":3217
            TabIndex        =   76
            Top             =   960
            Width           =   2895
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H0080FF80&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "transfer_cer.frx":33A3
            Left            =   6120
            List            =   "transfer_cer.frx":33B6
            TabIndex        =   74
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox Text24 
            BackColor       =   &H00FF8080&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   9600
            TabIndex        =   72
            Text            =   "Good"
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox Text23 
            BackColor       =   &H00FF8080&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6120
            TabIndex        =   70
            Text            =   "Satisfactory"
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox Text22 
            BackColor       =   &H00FF8080&
            Enabled         =   0   'False
            Height          =   405
            Left            =   2160
            TabIndex        =   68
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label35 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Admission Mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   8760
            TabIndex        =   79
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label34 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Special Remark"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   77
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label33 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Exam Appeared"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   75
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Result"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5280
            TabIndex        =   73
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Conduct"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   8760
            TabIndex        =   71
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label30 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Progress"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5280
            TabIndex        =   69
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inst. Leaving Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   360
            TabIndex        =   67
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "As per the request of candidate/parent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   65
         Top             =   5760
         Width           =   6255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Passed final year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   64
         Top             =   5280
         Width           =   4575
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Left            =   120
         TabIndex        =   51
         Top             =   6240
         Visible         =   0   'False
         Width           =   13095
         Begin VB.TextBox txtConduct 
            BackColor       =   &H00FF8080&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   9840
            TabIndex        =   5
            Text            =   "Good"
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H0080FF80&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "transfer_cer.frx":33FA
            Left            =   2040
            List            =   "transfer_cer.frx":3452
            TabIndex        =   8
            Top             =   1560
            Width           =   3015
         End
         Begin VB.TextBox txtRemark 
            BackColor       =   &H00FF8080&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   6840
            TabIndex        =   9
            Text            =   "AT HIS OWN REQUEST"
            Top             =   1680
            Width           =   3255
         End
         Begin VB.ComboBox cmbResult 
            BackColor       =   &H0080FF80&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "transfer_cer.frx":35DE
            Left            =   6240
            List            =   "transfer_cer.frx":35F1
            TabIndex        =   7
            Top             =   1080
            Width           =   2775
         End
         Begin VB.ComboBox CmbCourse 
            BackColor       =   &H0080FF80&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "transfer_cer.frx":3635
            Left            =   2040
            List            =   "transfer_cer.frx":364B
            TabIndex        =   6
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox Text21 
            BackColor       =   &H00FF8080&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6240
            TabIndex        =   4
            Text            =   "Satisfactory"
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox Text20 
            BackColor       =   &H00FF8080&
            Enabled         =   0   'False
            Height          =   405
            Left            =   2040
            TabIndex        =   3
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Current course"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   58
            Top             =   960
            Width           =   1275
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inst. Leaving Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   360
            TabIndex        =   57
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Conduct"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   9000
            TabIndex        =   56
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Progress"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5280
            TabIndex        =   55
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Result"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5280
            TabIndex        =   54
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Exam Appeared"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   53
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Special Remark"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5280
            TabIndex        =   52
            Top             =   1680
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3735
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   13095
         Begin VB.TextBox Text19 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5160
            TabIndex        =   50
            Top             =   3120
            Width           =   4695
         End
         Begin VB.TextBox Text18 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            TabIndex        =   48
            Top             =   2640
            Width           =   3375
         End
         Begin VB.TextBox Text17 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11040
            TabIndex        =   46
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7320
            TabIndex        =   44
            Top             =   2640
            Width           =   2175
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11040
            TabIndex        =   43
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7320
            TabIndex        =   42
            Top             =   2160
            Width           =   2175
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            TabIndex        =   41
            Top             =   2160
            Width           =   3375
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11040
            TabIndex        =   40
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7320
            TabIndex        =   39
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11040
            TabIndex        =   32
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7320
            TabIndex        =   31
            Top             =   1200
            Width           =   2175
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            TabIndex        =   30
            Top             =   1680
            Width           =   3375
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            TabIndex        =   29
            Top             =   1200
            Width           =   3375
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11040
            TabIndex        =   24
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11040
            TabIndex        =   23
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7320
            TabIndex        =   22
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7320
            TabIndex        =   21
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            TabIndex        =   20
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtRegNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            TabIndex        =   18
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Institution"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3735
            TabIndex        =   49
            Top             =   3120
            Width           =   1245
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contact No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   360
            TabIndex        =   47
            Top             =   2640
            Width           =   930
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Parent's No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9600
            TabIndex        =   45
            Top             =   2640
            Width           =   1035
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "DOB(words)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   360
            TabIndex        =   38
            Top             =   2160
            Width           =   1035
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Birth"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5640
            TabIndex        =   37
            Top             =   2160
            Width           =   1065
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Caste"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9600
            TabIndex        =   36
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Place "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9600
            TabIndex        =   35
            Top             =   2160
            Width           =   960
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Religion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5640
            TabIndex        =   34
            Top             =   2640
            Width           =   675
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5640
            TabIndex        =   33
            Top             =   1680
            Width           =   795
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Gender "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5640
            TabIndex        =   28
            Top             =   1200
            Width           =   705
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Mother's Name "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   360
            TabIndex        =   27
            Top             =   1680
            Width           =   1350
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Candidate Name "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   360
            TabIndex        =   26
            Top             =   1200
            Width           =   1470
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nationality"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9555
            TabIndex        =   25
            Top             =   1200
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Enrollment No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9600
            TabIndex        =   17
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Course  :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   16
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Admission year "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5640
            TabIndex        =   15
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Registration No. "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   360
            TabIndex        =   14
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Admission Date "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5640
            TabIndex        =   13
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Admission Mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9600
            TabIndex        =   12
            Top             =   240
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   13095
         Begin VB.TextBox Text1 
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   5160
            TabIndex        =   1
            Top             =   240
            Width           =   2775
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Search"
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8400
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Students Enrollment No. :-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1320
            TabIndex        =   19
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Label Label28 
         Caption         =   "Select following reason for student leaving certificate"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   63
         Top             =   4800
         Width           =   8415
      End
   End
End
Attribute VB_Name = "transfer_cer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim Conn As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Studdetail.mdb;Persist Security Info=False"
    Conn.Open
    Conn.Execute ("Insert Into TC_Details (EnrollNo, LeaveDate, Progress, Conduct, Course, Result, ExamGiven, Remarks, LeavDate, progress1, conduct1, ExamApper, result1, Remark, AdmitMode) Values ('" & Text1.Text & "','" & Text20.Text & "','" & Text21.Text & "','" & txtConduct.Text & "','" & CmbCourse.Text & "','" & cmbResult.Text & "','" & Combo1.Text & "','" & txtRemark.Text & "','" & Text22.Text & "','" & Text23.Text & "','" & Text24.Text & "','" & Combo3.Text & "','" & Combo2.Text & "','" & Text25.Text & "','" & Combo4.Text & "')")
    Conn.Close
    MsgBox "TC Data Generated Successfully"
    Text1.Text = ""
    Text1_Validate (False)
  
End Sub

Private Sub Command1_Click()
    Text20.Enabled = True
    Text21.Enabled = True
    txtConduct.Enabled = True
    CmbCourse.Enabled = True
    cmbResult.Enabled = True
    Combo1.Enabled = True
    txtRemark.Enabled = True
    Text22.Enabled = True
    Text23.Enabled = True
    Text24.Enabled = True
    Combo3.Enabled = True
    Combo2.Enabled = True
    Text25.Enabled = True
    Combo4.Enabled = True
    Command2.Enabled = True
    CmdSave.Enabled = True
    
    Dim Conn As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    Dim id As Integer
   
    Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Studdetail.mdb;Persist Security Info=False"
    Conn.Open
    Rs.Open "Select RegNo, AdmisDate, AdmisMode, Course, AdmisYear, EnrollNo, CandiName, Gender, Nationality, MomName, Category, Caste, DOBW, DOB, BirthPlace, ContactNo, Religion, Contact2, LastInst from Stud_reg where EnrollNo = '" + Text1.Text & "'", Conn
    
    txtRegNo.Text = Rs.Fields(0)
    Text3.Text = Rs.Fields(1)
    Text5.Text = Rs.Fields(2)
    Text2.Text = Rs.Fields(3)
    Text4.Text = Rs.Fields(4)
    Text6.Text = Rs.Fields(5)
    Text7.Text = Rs.Fields(6)
    Text9.Text = Rs.Fields(7)
    Text10.Text = Rs.Fields(8)
    Text8.Text = Rs.Fields(9)
    Text11.Text = Rs.Fields(10)
    Text12.Text = Rs.Fields(11)
    Text13.Text = Rs.Fields(12)
    Text14.Text = Rs.Fields(13)
    Text15.Text = Rs.Fields(14)
    Text18.Text = Rs.Fields(15)
    Text16.Text = Rs.Fields(16)
    Text17.Text = Rs.Fields(17)
    Text19.Text = Rs.Fields(18)
    Rs.Close
    Rs.Open "Select LeaveDate, Progress, Conduct, Course, Result, ExamGiven, Remarks, LeavDate, progress1, conduct1, ExamApper, result1, Remark, AdmitMode from TC_Details where EnrollNo = '" + Text1.Text & "'", Conn
    If (Not Rs.EOF) Then
        Text20.Text = Rs.Fields(0)
        Text21.Text = Rs.Fields(1)
        txtConduct.Text = Rs.Fields(2)
        CmbCourse.Text = Rs.Fields(3)
        cmbResult.Text = Rs.Fields(4)
        Combo1.Text = Rs.Fields(5)
        txtRemark.Text = Rs.Fields(6)
        Text22.Text = Rs.Fields(7)
        Text23.Text = Rs.Fields(8)
        Text24.Text = Rs.Fields(9)
        Combo3.Text = Rs.Fields(10)
        Combo2.Text = Rs.Fields(11)
        Text25.Text = Rs.Fields(12)
        Combo4.Text = Rs.Fields(13)
    End If
    Rs.Close
    Conn.Close
    Command2.SetFocus
End Sub

Private Sub Command2_Click()
    a = a + 1
    If DataEnvironment1.Connection1.State = 0 Then
        DataEnvironment1.Connection1.Open
    End If
    If a = 0 Then
    DataEnvironment1.Command1 Text1.Text, Text1.Text
    DataReport1.Show
    Else
    DataEnvironment1.Connection1.Close
    DataEnvironment1.Connection1.Open
    DataEnvironment1.Command1 Text1.Text, Text1.Text
    DataReport1.Show
    End If
    
    
End Sub


Private Sub Option1_Click()
If Option1.Value = True Then
Frame4.Visible = True
Frame6.Visible = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Frame6.Visible = True
Frame4.Visible = False
End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Text1.Text = "" Then
        Text1.SetFocus
        txtRegNo.Text = ""
        Text3.Text = ""
        Text5.Text = ""
        Text2.Text = ""
        Text4.Text = ""
        Text6.Text = ""
        Text7.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text8.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text18.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text19.Text = ""
        Text22.Enabled = False
        Text23.Enabled = False
        Text24.Enabled = False
        Combo3.Enabled = False
        Combo2.Enabled = False
        Text25.Enabled = False
        Text20.Enabled = False
        Text21.Enabled = False
        txtConduct.Enabled = False
        CmbCourse.Enabled = False
        cmbResult.Enabled = False
        Combo1.Enabled = False
        txtRemark.Enabled = False
        Combo4.Enabled = False
        Text20.Text = ""
        Text21.Text = "Satisfactory"
        txtConduct.Text = "Good"
        CmbCourse.Text = ""
        cmbResult.Text = ""
        Combo1.Text = ""
        txtRemark.Text = "AT HIS OWN REQUEST"
        Text22.Text = ""
        Text23.Text = "Satisfactory"
        Text24.Text = "Good"
        Combo2.Text = ""
        Combo3.Text = ""
        Combo4.Text = ""
        Text25.Text = "AT HIS OWN REQUEST"
        Command2.Enabled = False
        CmdSave.Enabled = False
    End If
End Sub

Private Sub Text20_LostFocus()
    If Not IsDate(Text20.Text) Then
        Text20.Text = ""
        MsgBox "Enter a valid date with this format: dd/mm/yyyy"
        Cancel = True
    End If
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 33 To 64, 91 To 96, 123 To 126
        MsgBox ("Must be a letter! Please try again!")
        Text21.Text = ""
        KeyAscii = 0
        Exit Sub
    End Select
End Sub

Private Sub txtConduct_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 33 To 64, 91 To 96, 123 To 126
        MsgBox ("Must be a letter! Please try again!")
        txtConduct.Text = ""
        KeyAscii = 0
        Exit Sub
    End Select
End Sub
