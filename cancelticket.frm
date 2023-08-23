VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form cancelticket 
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   8760
      Top             =   7320
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "CancelTicket"
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "Clear"
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "Submit"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "Select Date"
      Height          =   495
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   5160
      TabIndex        =   9
      Top             =   4200
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2021
      Month           =   3
      Day             =   30
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text4 
      DataField       =   "Mobie No"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   6360
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   5520
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      DataField       =   "Journey Date"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      DataField       =   "R no"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "NOTE: Tickets can only be canceled  4 days before the Journey Date "
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
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3360
      Width           =   7335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "Mobile No"
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
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "Name"
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
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Journey Date"
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
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
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
      Left            =   240
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   120
      Picture         =   "cancelticket.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   11655
   End
End
Attribute VB_Name = "cancelticket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = Calendar1.Value
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Fields("R no") = Text1.Text
Adodc1.Recordset.Fields("Journey Date") = Text2.Text
Adodc1.Recordset.Fields("Name") = Text3.Text
Adodc1.Recordset.Fields("Mobie No") = Text4.Text
Adodc1.Recordset.Update
MsgBox "Ticket Cancelation request submitted... ", vbInformation

End Sub

Private Sub Command3_Click()
menu.Show

End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
End Sub
