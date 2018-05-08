VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   12930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15555
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   12930
   ScaleWidth      =   15555
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   600
      Top             =   11520
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Suvajit\Documents\DATABASE.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Suvajit\Documents\DATABASE.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select distinct day_of_week month_of_year year from table1"
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
   Begin VB.CommandButton Command5 
      Caption         =   "GO"
      Height          =   375
      Left            =   14760
      TabIndex        =   16
      Top             =   12240
      Width           =   735
   End
   Begin VB.TextBox Text3 
      DataField       =   "gender"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Text            =   "Text3"
      Top             =   4560
      Width           =   3735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Suvajit\Documents\DATABASE.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Person"
      Top             =   12240
      Width           =   5895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CLEAR"
      Height          =   495
      Left            =   12720
      TabIndex        =   14
      Top             =   12120
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   10200
      TabIndex        =   13
      Top             =   12120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD"
      Height          =   495
      Left            =   7320
      TabIndex        =   12
      Top             =   12120
      Width           =   2295
   End
   Begin VB.ComboBox Combo4 
      DataField       =   "year"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   11
      Text            =   "Select"
      Top             =   9960
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "month_of_year"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   10
      Text            =   "Select"
      Top             =   8400
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "day_of_week"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   9
      Text            =   "Select"
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "occupation"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   3000
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      DataField       =   "name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1680
      Width           =   5775
   End
   Begin VB.Image ImageBetty 
      Height          =   3975
      Left            =   11760
      Picture         =   "form52.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image ImageRiya 
      Height          =   3975
      Left            =   11760
      Picture         =   "form52.frx":2155
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image ImageSumit 
      Height          =   3975
      Left            =   11760
      Picture         =   "form52.frx":AB47
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image ImageShreya 
      Height          =   3975
      Left            =   11760
      Picture         =   "form52.frx":10C7E
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "CHOOSE YEAR"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   9960
      Width           =   3375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "CHOOSE MONTH OF THE YEAR"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   5
      Top             =   8400
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "CHOOSE DAY OF THE WEEK"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   4
      Top             =   6480
      Width           =   3135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "GENDER"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   3
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "OCCUPATION"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   2
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TIME MANAGEMENT SYSTEM "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1095
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   10815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Private Sub Combo2_Load()
Combo2.AddItem "Sunday"
Combo2.AddItem "Monday"
Combo2.AddItem "Tuesday"
Combo2.AddItem "Wednesday"
Combo2.AddItem "Thursday"
Combo2.AddItem "Friday"
Combo2.AddItem "Saturday"

End Sub











Private Sub Combo3_Change()
Combo3.AddItem "January"
Combo3.AddItem "February"
Combo3.AddItem "March"
Combo3.AddItem "April"
Combo3.AddItem "May"
Combo3.AddItem "June"
Combo3.AddItem "July"
Combo3.AddItem "August"
Combo3.AddItem "September"
Combo3.AddItem "October"
Combo3.AddItem "November"
Combo3.AddItem "December"

End Sub



Private Sub Combo4_Change()
Combo4.Addtem "1997"
Combo4.Addtem "1998"
Combo4.Addtem "1999"
Combo4.Addtem "2000"
Combo4.Addtem "2001"
Combo4.Addtem "2002"
Combo4.Addtem "2003"
Combo4.Addtem "2004"
Combo4.Addtem "2005"
Combo4.Addtem "2006"
Combo4.Addtem "2007"
Combo4.Addtem "2008"
Combo4.Addtem "2009"
Combo4.Addtem "2010"
Combo4.Addtem "2011"
Combo4.Addtem "2012"
Combo4.Addtem "2013"
Combo4.Addtem "2014"
Combo4.Addtem "2015"
Combo4.Addtem "2016"
Combo4.Addtem "2017"
Combo4.Addtem "2018"
Combo4.Addtem "2019"
Combo4.Addtem "2020"


End Sub

Private Sub Command2_Click()
Data1.Recordset.AddNew


End Sub

Private Sub Command3_Click()
Data1.Recordset.Update
Data1.Recordset.Requery
Data1.Refresh




End Sub

Private Sub Command4_Click()
Text1.Text = " "
Text2.Text = " "

End Sub


Private Sub Command5_Click()
Form1.Show

End Sub





Private Sub Text1_Change()
Call invisible

If Text1.Text = "Shreya" Then
ImageShreya.Visible = True
ElseIf Text1.Text = "Sumit" Then
ImageSumit.Visible = True
ElseIf Text1.Text = "Riya" Then
ImageRiya.Visible = True
ElseIf Text1.Text = "Betty" Then
ImageBetty.Visible = True
End If



End Sub
 Private Sub invisible()
 ImageShreya.Visible = False
 ImageSumit.Visible = False
 ImageRiya.Visible = False
 ImageBetty.Visible = False
 End Sub

