VERSION 5.00
Begin VB.Form student_reg 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Student Details"
   ClientHeight    =   10800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   Icon            =   "student_reg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   13680
      ScaleHeight     =   270
      ScaleWidth      =   1140
      TabIndex        =   48
      Top             =   8400
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   8055
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   14415
      Begin VB.Frame Frame4 
         Height          =   4815
         Left            =   240
         TabIndex        =   25
         Top             =   2160
         Width           =   13935
         Begin VB.ComboBox comReligion 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            ItemData        =   "student_reg.frx":0CCA
            Left            =   1800
            List            =   "student_reg.frx":0CE3
            TabIndex        =   16
            Text            =   "SELECT RELIGION"
            Top             =   2040
            Width           =   2415
         End
         Begin VB.Frame Frame5 
            Height          =   495
            Left            =   8040
            TabIndex        =   47
            Top             =   240
            Width           =   2895
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFC0FF&
               Caption         =   "Male"
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
               Left            =   0
               TabIndex        =   9
               Top             =   120
               Width           =   1455
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00FFC0FF&
               Caption         =   "Female"
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
               Left            =   1440
               TabIndex        =   10
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.ComboBox comCategory 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            ItemData        =   "student_reg.frx":0D1E
            Left            =   5760
            List            =   "student_reg.frx":0D2E
            TabIndex        =   17
            Text            =   "SELECT CATEGORY"
            Top             =   2040
            Width           =   1935
         End
         Begin VB.TextBox DOBW 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   9360
            TabIndex        =   15
            Top             =   1320
            Width           =   4455
         End
         Begin VB.TextBox DOB 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5760
            TabIndex        =   14
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox Nation 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8160
            TabIndex        =   12
            Text            =   "INDIAN"
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Contact2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5760
            TabIndex        =   20
            Top             =   2520
            Width           =   1695
         End
         Begin VB.TextBox Email 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8520
            TabIndex        =   21
            Top             =   2520
            Width           =   5295
         End
         Begin VB.TextBox Addr 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Left            =   1800
            TabIndex        =   22
            Top             =   3000
            Width           =   4935
         End
         Begin VB.TextBox LastInst 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1800
            TabIndex        =   23
            Top             =   4080
            Width           =   6135
         End
         Begin VB.TextBox ContactNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   19
            Top             =   2520
            Width           =   2415
         End
         Begin VB.TextBox Caste 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8520
            TabIndex        =   18
            Top             =   2040
            Width           =   3255
         End
         Begin VB.TextBox BirthPlace 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1800
            TabIndex        =   13
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox CandiName 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   8
            Top             =   360
            Width           =   4695
         End
         Begin VB.TextBox MomName 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   11
            Top             =   840
            Width           =   4695
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6840
            TabIndex        =   40
            Top             =   840
            Width           =   1065
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4440
            TabIndex        =   39
            Top             =   2040
            Width           =   915
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Parent's No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   4440
            TabIndex        =   38
            Top             =   2520
            Width           =   1185
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Email ID "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7560
            TabIndex        =   37
            Top             =   2520
            Width           =   810
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   36
            Top             =   4200
            Width           =   1500
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   35
            Top             =   3120
            Width           =   1335
            WordWrap        =   -1  'True
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   34
            Top             =   2520
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Candidate Name "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1620
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   32
            Top             =   2040
            Width           =   750
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   1095
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   7800
            TabIndex        =   30
            Top             =   2040
            Width           =   585
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   4320
            TabIndex        =   29
            Top             =   1440
            Width           =   1245
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   1515
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   8040
            TabIndex        =   27
            Top             =   1440
            Width           =   1170
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   7200
            TabIndex        =   26
            Top             =   360
            Width           =   765
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000A&
         Height          =   1095
         Left            =   0
         TabIndex        =   24
         Top             =   6960
         Width           =   14415
         Begin VB.CommandButton Command2 
            BackColor       =   &H8000000E&
            Caption         =   "&Save"
            Default         =   -1  'True
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
            Left            =   6720
            Picture         =   "student_reg.frx":0D47
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Save the Data"
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton Command3 
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
            Left            =   8040
            Picture         =   "student_reg.frx":1989
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "close this window"
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H8000000E&
            Caption         =   "&Reset"
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
            Left            =   5280
            Picture         =   "student_reg.frx":25CB
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Reset current item"
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H8000000E&
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1560
            Picture         =   "student_reg.frx":320D
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "close this window"
            Top             =   1920
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Student Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   13935
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0E0FF&
            Height          =   375
            Left            =   6960
            TabIndex        =   56
            Text            =   "2016"
            Top             =   480
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox tcNo 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox EnrollNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10560
            TabIndex        =   7
            Top             =   1560
            Width           =   2775
         End
         Begin VB.ComboBox comCourse 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "student_reg.frx":3E4F
            Left            =   1200
            List            =   "student_reg.frx":3E65
            TabIndex        =   5
            Text            =   "SELECT COURSE"
            Top             =   1560
            Width           =   4095
         End
         Begin VB.ComboBox comAdmisYear 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "student_reg.frx":3F0D
            Left            =   6960
            List            =   "student_reg.frx":3F50
            TabIndex        =   6
            Text            =   "SELECT YEAR"
            Top             =   1560
            Width           =   2055
         End
         Begin VB.TextBox AdmisDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            TabIndex        =   3
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox RegiNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            TabIndex        =   1
            Top             =   960
            Width           =   3015
         End
         Begin VB.ComboBox comAdmisMode 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "student_reg.frx":3FD2
            Left            =   10800
            List            =   "student_reg.frx":3FDC
            TabIndex        =   4
            Text            =   "SELECT MODE"
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label17 
            Caption         =   "INST. REG. No. "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5400
            TabIndex        =   55
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "TC No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   480
            Width           =   1335
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9120
            TabIndex        =   46
            Top             =   960
            Width           =   1560
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5280
            TabIndex        =   45
            Top             =   960
            Width           =   1560
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Registration No. "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   1650
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5400
            TabIndex        =   43
            Top             =   1560
            Width           =   1545
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   42
            Top             =   1560
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Enrollment No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9120
            TabIndex        =   41
            Top             =   1560
            Width           =   1320
         End
      End
   End
End
Attribute VB_Name = "student_reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AdmisDate_Validate(Cancel As Boolean)
    If Not IsDate(AdmisDate.Text) Then
        AdmisDate.Text = ""
        MsgBox "Enter a valid date with this format: dd/mm/yyyy"
        Cancel = True
    End If
End Sub

Private Sub BirthPlace_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 33 To 64, 91 To 96, 123 To 126
        MsgBox ("Must be a letter! Please try again!")
        BirthPlace.Text = ""
        KeyAscii = 0
        Exit Sub
    End Select
End Sub

Private Sub CandiName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 33 To 64, 91 To 96, 123 To 126
        MsgBox ("Must be a letter! Please try again!")
        CandiName.Text = ""
        KeyAscii = 0
        Exit Sub
    End Select
      Select Case KeyAscii1
    Case 97 To 126
        MsgBox ("Enter Capital alphabets")
        CandiName.Text = ""
        KeyAscii1 = 0
        Exit Sub
    End Select
End Sub

Private Sub Caste_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 33 To 64, 91 To 96, 123 To 126
        MsgBox ("Must be a letter! Please try again!")
        Caste.Text = ""
        KeyAscii = 0
        Exit Sub
    End Select
End Sub

Private Sub Command2_Click()
    If tcNo.Text = "" Or RegiNo.Text = "" Or AdmisDate.Text = "" Or comAdmisMode = "" Or comCourse.Text = "" Or comAdmisYear.Text = "" Or EnrollNo.Text = "" Or CandiName.Text = "" Or MomName.Text = "" Or Nation.Text = "" Or BirthPlace.Text = "" Or DOB.Text = "" Or DOBW.Text = "" Or comReligion.Text = "" Or comCategory.Text = "" Or Caste.Text = "" Or ContactNo.Text = "" Or Contact2.Text = "" Or Email.Text = "" Or Addr.Text = "" Or LastInst.Text = "" Or Text1.Text = "" Then
    MsgBox "Filled the all Record to procced"
    Else
    Dim Conn As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    Dim id As Integer
    Dim gen As String
    If Option1.Value = True Then
        gen = "Male"
    Else
        gen = "Female"
    End If
    Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Studdetail.mdb;Persist Security Info=False"
    Conn.Open
    Conn.Execute ("Insert Into Stud_reg (AdmisDate, AdmisMode, Course, AdmisYear, EnrollNo, CandiName, Gender, MomName, Nationality, BirthPlace, DOB, DOBW, Religion, Category, Caste, ContactNo, Contact2, Email, Address, LastInst, RegNo, INST_REG_NO ) Values ('" & AdmisDate.Text & "','" & comAdmisMode.Text & "','" & comCourse.Text & "','" & comAdmisYear.Text & "','" & EnrollNo.Text & "','" & CandiName.Text & "','" & gen & "','" & MomName.Text & "','" & Nation.Text & "','" & BirthPlace.Text & "','" & DOB.Text & "','" & DOBW.Text & "','" & comReligion.Text & "','" & comCategory.Text & "','" & Caste.Text & "','" & ContactNo.Text & "','" & Contact2.Text & "','" & Email.Text & "','" & Addr.Text & "','" & LastInst.Text & "','" & RegiNo.Text & "','" & Text1.Text & "')")
    Conn.Close
    MsgBox "Student Data Generated Successfully"
    tcNo.Text = ""
    RegiNo.Text = ""
    AdmisDate.Text = ""
    comAdmisMode = "SELECT MODE"
    comCourse.Text = "SELECT COURSE"
    comAdmisYear.Text = "SELECT YEAR"
    EnrollNo.Text = ""
    CandiName.Text = ""
    MomName.Text = ""
    Nation.Text = "INDIAN"
    BirthPlace.Text = ""
    DOB.Text = ""
    DOBW.Text = ""
    comReligion.Text = "SELECT RELIGION"
    comCategory.Text = "SELECT CATEGORY"
    Caste.Text = ""
    ContactNo.Text = ""
    Contact2.Text = ""
    Email.Text = ""
    Addr.Text = ""
    LastInst.Text = ""
    Text1.Text = ""
    Option1.Value = False
    Option2.Value = False
    Module1.NewTCNo
 End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
RegiNo.Text = ""
AdmisDate.Text = ""
comAdmisMode = "SELECT MODE"
comCourse.Text = "SELECT COURSE"
comAdmisYear.Text = "SELECT YEAR"
EnrollNo.Text = ""
CandiName.Text = ""
MomName.Text = ""
Nation.Text = "INDIAN"
BirthPlace.Text = ""
DOB.Text = ""
DOBW.Text = ""
comReligion.Text = "SELECT RELIGION"
comCategory.Text = "SELECT CATEGORY"
Caste.Text = ""
ContactNo.Text = ""
Contact2.Text = ""
Email.Text = ""
Addr.Text = ""
LastInst.Text = ""
Text1.Text = ""
Option1.Value = False
Option2.Value = False
End Sub





Private Sub DOB_Validate(Cancel As Boolean)
    If Not IsDate(DOB.Text) Then
        DOB.Text = ""
        MsgBox "Enter a valid date with this format: dd/mm/yyyy"
        Cancel = True
    End If
End Sub




Private Sub EnrollNo_Validate(Cancel As Boolean)
    If Not IsNumeric(EnrollNo.Text) Then
      EnrollNo = ""
        MsgBox "Enter only Numeric Values"
       Cancel = True
       EnrollNo.SetFocus
        ElseIf Len(EnrollNo.Text) < 10 Then
      EnrollNo = ""
      MsgBox "Enrollment Number is too Short"
       Cancel = True
      ContactNo.SetFocus
    ElseIf Len(EnrollNo.Text) > 10 Then
    EnrollNo = ""
    MsgBox "Enrollment Number is too Long"
     Cancel = True
    EnrollNo.SetFocus
   End If
End Sub

Private Sub Form_Load()
    Module1.NewTCNo
End Sub

Private Sub MomName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 33 To 64, 91 To 96, 123 To 126
        MsgBox ("Must be a letter! Please try again!")
        MomName.Text = ""
        KeyAscii = 0
        Exit Sub
    End Select
End Sub

Private Sub Nation_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 33 To 64, 91 To 96, 123 To 126
        MsgBox ("Must be a letter! Please try again!")
        Nation.Text = ""
        KeyAscii = 0
        Exit Sub
    End Select
End Sub


