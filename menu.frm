VERSION 5.00
Begin VB.Form menu 
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C000&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C000&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "Bus Services"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "Customer care"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "Cancel Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "Book Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   120
      Picture         =   "menu.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
bookticket.Show



End Sub

Private Sub Command2_Click()
cancelticket.Show

End Sub

Private Sub Command3_Click()
query.Show
End Sub

Private Sub Command4_Click()
busservices.Show

End Sub

Private Sub Command5_Click()
login.Show
menu.Hide

End Sub

Private Sub Command6_Click()
Priticket.Show




End Sub
