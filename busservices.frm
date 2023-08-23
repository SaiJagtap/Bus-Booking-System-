VERSION 5.00
Begin VB.Form busservices 
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8040
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   7440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   4440
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4440
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Caption         =   "Luxury Bus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3855
      Left            =   7320
      TabIndex        =   2
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Semi-Luxury Bus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3855
      Left            =   3840
      TabIndex        =   1
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Normal Bus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   120
      Picture         =   "busservices.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   10815
   End
End
Attribute VB_Name = "busservices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
menu.Show
busservices.Hide

End Sub

Private Sub Form_Load()
Text1.Text = "Comfortable Seats with proper Infrastructure.Price are affordable and prices ranges between Rs200-1500 (Depends upon your Destination).Hygienity maintain."
Text2.Text = "Comfortable Seats with proper Infrastructure.Price are affordable and prices ranges between Rs500-3500 (Depends upon your Destination).Hygienity maintain,T.V and Wifi facility available."
Text3.Text = "A/C Seats with proper Infrastructure.Price are affordable and prices ranges between Rs1000-10,000 (Depends upon your Destination).Hygienity maintain,T.V and Wifi facility available."

              

End Sub

