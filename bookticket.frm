VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form bookticket 
   Caption         =   "Form1"
   ClientHeight    =   9900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      BackColor       =   &H8000000B&
      Caption         =   "Conform"
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8760
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      DataField       =   "Price"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1920
      TabIndex        =   24
      Top             =   8760
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   8640
      Top             =   8400
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
      RecordSource    =   "BookTicket"
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
   Begin VB.TextBox Text7 
      DataField       =   "E-mail"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   6840
      TabIndex        =   22
      Top             =   5400
      Width           =   3255
   End
   Begin VB.TextBox Text6 
      DataField       =   "Mobile No"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   6840
      TabIndex        =   21
      Top             =   4560
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   6840
      TabIndex        =   20
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
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
      TabIndex        =   16
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9360
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9360
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "Bus Type"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1920
      TabIndex        =   13
      Top             =   8040
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Gender"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1920
      TabIndex        =   12
      Top             =   7200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Date"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Width           =   1215
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2295
      Left            =   4920
      TabIndex        =   10
      Top             =   6240
      Width           =   3615
      _Version        =   524288
      _ExtentX        =   6376
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2021
      Month           =   3
      Day             =   29
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
      DataField       =   "Date of Journey"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   6240
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      DataField       =   "Destination Place"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      DataField       =   "Pickup Place"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      DataField       =   "R no"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C000&
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
      Left            =   120
      TabIndex        =   23
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C000&
      Caption         =   "E-mail"
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
      Left            =   5400
      TabIndex        =   19
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label8 
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
      Left            =   5400
      TabIndex        =   18
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label7 
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
      Left            =   5400
      TabIndex        =   17
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C000&
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
      Left            =   120
      TabIndex        =   5
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      Caption         =   "Gender"
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
      TabIndex        =   4
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
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
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
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
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   1455
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
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   120
      Picture         =   "bookticket.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   11535
   End
End
Attribute VB_Name = "bookticket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text4.Text = Calendar1.Value
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Fields("R no") = Text1.Text
Adodc1.Recordset.Fields("Pickup Place") = Text2.Text
Adodc1.Recordset.Fields("Destination Place") = Text3.Text
Adodc1.Recordset.Fields("Date of Journey") = Text4.Text
Adodc1.Recordset.Fields("Gender") = Combo1.Text
Adodc1.Recordset.Fields("Bus Type") = Combo2.Text
Adodc1.Recordset.Fields("Name") = Text5.Text
Adodc1.Recordset.Fields("Mobile No") = Text6.Text
Adodc1.Recordset.Fields("E-mail") = Text7.Text
Adodc1.Recordset.Fields("Price") = Text8.Text

Adodc1.Recordset.Update
MsgBox "Your Ticket has been booked", vbInformation
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Combo1.Text = ""
Combo2.Text = ""



End Sub

Private Sub Command4_Click()
menu.Show

End Sub

Private Sub Command5_Click()
If Text2.Text = "mumbai" And Text3.Text = "pune" And Combo2.Text = "Normal Bus" Then
Text8.Text = "350"
Else
If Text2.Text = "pune" And Text3.Text = "mumbai" And Combo2.Text = "Normal Bus" Then
Text8.Text = "350"
Else
If Text2.Text = "mumbai" And Text3.Text = "satara" And Combo2.Text = "Normal Bus" Then
Text8.Text = "500"
Else
If Text2.Text = "satara" And Text3.Text = "mumbai" And Combo2.Text = "Normal Bus" Then
Text8.Text = "500"
Else
If Text2.Text = "pune" And Text3.Text = "satara" And Combo2.Text = "Normal Bus" Then
Text8.Text = "245"
Else
If Text2.Text = "satara" And Text3.Text = "pune" And Combo2.Text = "Normal Bus" Then
Text8.Text = "245"
Else
If Text2.Text = "mumbai" And Text3.Text = "nanded" And Combo2.Text = "Normal Bus" Then
Text8.Text = "800"
Else
If Text2.Text = "nanded" And Text3.Text = "mumbai" And Combo2.Text = "Normal Bus" Then
Text8.Text = "800"
Else
If Text2.Text = "nanded" And Text3.Text = "pune" And Combo2.Text = "Normal Bus" Then
Text8.Text = "750"
Else
If Text2.Text = "pune" And Text3.Text = "nanded" And Combo2.Text = "Normal Bus" Then
Text8.Text = "750"
Else
If Text2.Text = "satara" And Text3.Text = "nanded" And Combo2.Text = "Normal Bus" Then
Text8.Text = "500"
Else
If Text2.Text = "nanded" And Text3.Text = "satara" And Combo2.Text = "Normal Bus" Then
Text8.Text = "500"
Else
If Text2.Text = "mumbai" And Text3.Text = "kholapur" And Combo2.Text = "Normal Bus" Then
Text8.Text = "900"
Else
If Text2.Text = "kholapur" And Text3.Text = "mumbai" And Combo2.Text = "Normal Bus" Then
Text8.Text = "900"
Else
If Text2.Text = "pune" And Text3.Text = "kholapur" And Combo2.Text = "Normal Bus" Then
Text8.Text = "650"
Else
If Text2.Text = "kholapur" And Text3.Text = "pune" And Combo2.Text = "Normal Bus" Then
Text8.Text = "650"
Else
If Text2.Text = "satara" And Text3.Text = "kholapur" And Combo2.Text = "Normal Bus" Then
Text8.Text = "300"
Else
If Text2.Text = "kholapur" And Text3.Text = "satara" And Combo2.Text = "Normal Bus" Then
Text8.Text = "300"
Else
If Text2.Text = "pune" And Text3.Text = "goa" And Combo2.Text = "Normal Bus" Then
Text8.Text = "1000"
Else
If Text2.Text = "goa" And Text3.Text = "pune" And Combo2.Text = "Normal Bus" Then
Text8.Text = "1000"
Else
If Text2.Text = "mumbai" And Text3.Text = "goa" And Combo2.Text = "Normal Bus" Then
Text8.Text = "1500"
Else
If Text2.Text = "goa" And Text3.Text = "mumbai" And Combo2.Text = "Normal Bus" Then
Text8.Text = "1500"
Else
If Text2.Text = "satara" And Text3.Text = "goa" And Combo2.Text = "Normal Bus" Then
Text8.Text = "1350"
Else
If Text2.Text = "goa" And Text3.Text = "satara" And Combo2.Text = "Normal Bus" Then
Text8.Text = "1350"
Else
If Text2.Text = "mumbai" And Text3.Text = "pune" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "650"
Else
If Text2.Text = "pune" And Text3.Text = "mumbai" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "650"
Else
If Text2.Text = "mumbai" And Text3.Text = "satara" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "1000"
Else
If Text2.Text = "satara" And Text3.Text = "mumbai" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "1000"
Else
If Text2.Text = "pune" And Text3.Text = "satara" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "600"
Else
If Text2.Text = "satara" And Text3.Text = "pune" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "600"
Else
If Text2.Text = "mumbai" And Text3.Text = "nanded" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "2000"
Else
If Text2.Text = "nanded" And Text3.Text = "mumbai" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "2000"
Else
If Text2.Text = "nanded" And Text3.Text = "pune" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "1450"
Else
If Text2.Text = "pune" And Text3.Text = "nanded" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "1450"
Else
If Text2.Text = "satara" And Text3.Text = "nanded" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "1000"
Else
If Text2.Text = "nanded" And Text3.Text = "satara" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "1000"
Else
If Text2.Text = "mumbai" And Text3.Text = "kholapur" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "1200"
Else
If Text2.Text = "kholapur" And Text3.Text = "mumbai" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "1200"
Else
If Text2.Text = "pune" And Text3.Text = "kholapur" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "950"
Else
If Text2.Text = "kholapur" And Text3.Text = "pune" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "950"
Else
If Text2.Text = "satara" And Text3.Text = "kholapur" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "750"
Else
If Text2.Text = "kholapur" And Text3.Text = "satara" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "750"
Else
If Text2.Text = "pune" And Text3.Text = "goa" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "1500"
Else
If Text2.Text = "goa" And Text3.Text = "pune" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "1500"
Else
If Text2.Text = "mumbai" And Text3.Text = "goa" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "2800"
Else
If Text2.Text = "goa" And Text3.Text = "mumbai" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "2800"
Else
If Text2.Text = "satara" And Text3.Text = "goa" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "1900"
Else
If Text2.Text = "goa" And Text3.Text = "satara" And Combo2.Text = "Semi-Luxury Bus" Then
Text8.Text = "1900"
Else
If Text2.Text = "mumbai" And Text3.Text = "pune" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "4500"
Else
If Text2.Text = "pune" And Text3.Text = "mumbai" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "4500"
Else
If Text2.Text = "mumbai" And Text3.Text = "satara" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "6000"
Else
If Text2.Text = "satara" And Text3.Text = "mumbai" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "6000"
Else
If Text2.Text = "pune" And Text3.Text = "satara" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "2300"
Else
If Text2.Text = "satara" And Text3.Text = "pune" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "2300"
Else
If Text2.Text = "mumbai" And Text3.Text = "nanded" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "5000"
Else
If Text2.Text = "nanded" And Text3.Text = "mumbai" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "5000"
Else
If Text2.Text = "nanded" And Text3.Text = "pune" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "4500"
Else
If Text2.Text = "pune" And Text3.Text = "nanded" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "4500"
Else
If Text2.Text = "satara" And Text3.Text = "nanded" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "2000"
Else
If Text2.Text = "nanded" And Text3.Text = "satara" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "2000"
Else
If Text2.Text = "mumbai" And Text3.Text = "kholapur" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "3500"
Else
If Text2.Text = "kholapur" And Text3.Text = "mumbai" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "3500"
Else
If Text2.Text = "pune" And Text3.Text = "kholapur" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "2500"
Else
If Text2.Text = "kholapur" And Text3.Text = "pune" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "2500"
Else
If Text2.Text = "satara" And Text3.Text = "kholapur" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "1900"
Else
If Text2.Text = "kholapur" And Text3.Text = "satara" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "1900"
Else
If Text2.Text = "pune" And Text3.Text = "goa" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "3900"
Else
If Text2.Text = "goa" And Text3.Text = "pune" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "3900"
Else
If Text2.Text = "mumbai" And Text3.Text = "goa" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "4500"
Else
If Text2.Text = "goa" And Text3.Text = "mumbai" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "4500"
Else
If Text2.Text = "satara" And Text3.Text = "goa" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "2500"
Else
If Text2.Text = "goa" And Text3.Text = "satara" And Combo2.Text = "Luxury Bus" Then
Text8.Text = "2500"
Else
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Form_Load()
Combo1.AddItem ("Male")
Combo1.AddItem ("Female")
Combo2.AddItem ("Normal Bus")
Combo2.AddItem ("Semi-Luxury Bus")
Combo2.AddItem ("Luxury Bus")
Adodc1.Recordset.AddNew




End Sub

