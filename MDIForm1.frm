VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Student Leaving Certificate"
   ClientHeight    =   10650
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   20250
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10275
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "File "
      Begin VB.Menu sr 
         Caption         =   "Student Registration"
         Checked         =   -1  'True
      End
      Begin VB.Menu tc 
         Caption         =   "Transfer certificate"
         Checked         =   -1  'True
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
Form4.Show
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub MDIForm_Load()
Dim a As Panel
StatusBar1.Panels(1).Text = "Welcome GPW"
Set a = StatusBar1.Panels.Add(, , , sbrDate)
Set a = StatusBar1.Panels.Add(, , , sbrTime)
End Sub

Private Sub sr_Click()
student_reg.Show
End Sub



Private Sub tc_Click()
transfer_cer.Show
End Sub
