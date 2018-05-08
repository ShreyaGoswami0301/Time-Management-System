VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form3"
   ClientHeight    =   12930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15555
   LinkTopic       =   "Form3"
   ScaleHeight     =   12930
   ScaleWidth      =   15555
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   480
      Top             =   11400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DATABASE.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DATABASE.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select from DATABASE"
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
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   11520
      ScaleHeight     =   3435
      ScaleWidth      =   3195
      TabIndex        =   21
      Top             =   1800
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "GO"
      Height          =   375
      Left            =   14760
      TabIndex        =   20
      Top             =   12240
      Width           =   735
   End
   Begin VB.TextBox Text3 
      DataField       =   "gender"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4920
      TabIndex        =   19
      Text            =   "Text3"
      Top             =   5520
      Width           =   3735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "FEMALE"
      Height          =   375
      Left            =   6240
      TabIndex        =   18
      Top             =   4800
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MALE"
      Height          =   255
      Left            =   4800
      TabIndex        =   17
      Top             =   4800
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Suvajit\Documents\DATABASE.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
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
      TabIndex        =   16
      Top             =   12120
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   10200
      TabIndex        =   15
      Top             =   12120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD"
      Height          =   495
      Left            =   7320
      TabIndex        =   14
      Top             =   12120
      Width           =   2295
   End
   Begin VB.OptionButton Option3 
      Caption         =   "OTHERS"
      Height          =   375
      Left            =   7560
      TabIndex        =   13
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPLOAD PHOTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      TabIndex        =   12
      Top             =   5520
      Width           =   2895
   End
   Begin VB.ComboBox Combo4 
      DataField       =   "year"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   4920
      TabIndex        =   11
      Text            =   "Combo4"
      Top             =   9960
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "month_of_year"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   4920
      TabIndex        =   10
      Text            =   "Combo3"
      Top             =   8400
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "day_of_week"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   4920
      TabIndex        =   9
      Text            =   "Combo2"
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "occupation"
      DataSource      =   "Data1"
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
      Height          =   855
      Left            =   4680
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1680
      Width           =   5775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "CHOOSE YEAR"
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   9960
      Width           =   3375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "CHOOSE MONTH OF THE YEAR"
      Height          =   1095
      Left            =   480
      TabIndex        =   5
      Top             =   8400
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "CHOOSE DAY OF THE WEEK"
      Height          =   1095
      Left            =   600
      TabIndex        =   4
      Top             =   6480
      Width           =   3135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "GENDER"
      Height          =   1095
      Left            =   720
      TabIndex        =   3
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "OCCUPATION"
      Height          =   855
      Left            =   720
      TabIndex        =   2
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "NAME"
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
Private Sub Combo2_Change()
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
Combo4.AddItem "1997"
 Combo4.AddItem "1998"
 Combo4.AddItem "1999"
 Combo4.AddItem "2000"
 Combo4.AddItem "2001"
 Combo4.AddItem "2002"
 Combo4.AddItem "2003"
 Combo4.AddItem "2004"
 Combo4.AddItem "2005"
 Combo4.AddItem "2006"
 Combo4.AddItem "2007"
 Combo4.AddItem "2008"
 Combo4.AddItem "2009"
 Combo4.AddItem "2010"
 Combo4.AddItem "2011"
 Combo4.AddItem "2012"
 Combo4.AddItem "2013"
 Combo4.AddItem "2014"
 Combo4.AddItem "2015"
 Combo4.AddItem "2016"
 Combo4.AddItem "2017"
 Combo4.AddItem "2018"
 Combo4.AddItem "2019"
 Combo4.AddItem "2020"
End Sub

Private Sub Command1_Click()
commondialog1.FileName = " "


End Sub

Private Sub Command2_Click()
Data1.Recordset.AddNew


End Sub

Private Sub Command3_Click()
Data1.Recordset.Update

End Sub

Private Sub Command4_Click()
Text1.Text = " "
Text2.Text = " "

End Sub


Private Sub Command5_Click()
Form12.Show

End Sub

Private Sub Form_Load()
If Option1.Value = True Then
Gender = "Male"
ElseIf Option2.Value = True Then
Gender = "Female"
ElseIf Option3.Value = True Then
Gender = "Others"
Text3.Text = Gender

End If


End Sub
