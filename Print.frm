VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Priticket 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   4680
      Top             =   6240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Sunil Jagtap\Desktop\5semproject\signin.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Sunil Jagtap\Desktop\5semproject\signin.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from BookTicket"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   7455
      Begin VB.TextBox Text6 
         DataField       =   "Price"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1920
         TabIndex        =   12
         Top             =   4440
         Width           =   3735
      End
      Begin VB.TextBox Text5 
         DataField       =   "Bus Type"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1920
         TabIndex        =   11
         Top             =   3600
         Width           =   3735
      End
      Begin VB.TextBox Text4 
         DataField       =   "Date of Journey"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1920
         TabIndex        =   10
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         DataField       =   "Destination Place"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1920
         TabIndex        =   9
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         DataField       =   "Pickup Place"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1920
         TabIndex        =   8
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FFFF&
         Caption         =   "H    A    P    P    Y    J    O    U    R    N    E    Y"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   3495
         Left            =   6360
         TabIndex        =   16
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FFFF&
         Caption         =   "J's Bus Travels"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   14
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Price"
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
         Left            =   240
         TabIndex        =   7
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Bus Type"
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
         Left            =   240
         TabIndex        =   6
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Date of Journey"
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
         Left            =   240
         TabIndex        =   5
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Destination Place"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Pickup Place"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Enter the correct registration number(R.no) of your booked ticket.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "R no"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Priticket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

 
 
 
 
 End If
 





End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""


End Sub
