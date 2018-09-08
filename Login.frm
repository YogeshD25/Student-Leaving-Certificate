VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   Caption         =   "Student Leaving Certificate"
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H8000000E&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4800
         Picture         =   "Login.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "close this window"
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H8000000E&
         Caption         =   "LOGIN"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         Picture         =   "Login.frx":190C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Login your account"
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox nm 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   2
         Text            =   "Student Section "
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox pass 
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   4560
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF80&
         Caption         =   "WELCOME TO GOVERNMENT POLYTECHNIC WASHIM"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   9735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF80&
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF80&
         Caption         =   "Student Leaving Certificate"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   2280
         TabIndex        =   4
         Top             =   1320
         Width           =   5415
      End
      Begin VB.Image img_AddNewEmployee 
         Height          =   480
         Left            =   1560
         Picture         =   "Login.frx":254E
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF80&
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   3
         Top             =   3000
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub CmdCancel_Click()
    LoginSucceeded = False
    End
End Sub

Private Sub cmdOK_Click()
    If pass = "sakhare" Then
        LoginSucceeded = True
        Unload Me
        MDIForm1.Show
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        pass.SetFocus
        pass.Text = ""
    End If
End Sub

