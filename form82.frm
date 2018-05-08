VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   5880
      Top             =   10680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "form82.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5760
      Top             =   9120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "form82.frx":72C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "form82.frx":921C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text6 
      Height          =   480
      Left            =   10320
      TabIndex        =   19
      Top             =   9960
      Width           =   4335
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   10320
      TabIndex        =   18
      Top             =   8400
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      Height          =   465
      Left            =   10320
      TabIndex        =   17
      Top             =   6480
      Width           =   4335
   End
   Begin MSComctlLib.ImageCombo ImageCombo3 
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   9960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "Select"
   End
   Begin MSComctlLib.ImageCombo ImageCombo2 
      Height          =   495
      Left            =   4680
      TabIndex        =   15
      Top             =   8520
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "Select"
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Top             =   6480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "Select"
   End
   Begin VB.CommandButton Command5 
      Caption         =   "GO"
      Height          =   375
      Left            =   14760
      TabIndex        =   13
      Top             =   12240
      Width           =   735
   End
   Begin VB.TextBox Text3 
      DataField       =   "gender"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Text            =   "Enter gender"
      Top             =   4560
      Width           =   5655
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Suvajit\Desktop\project vb 3.0\DATABASE.mdb"
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
      TabIndex        =   11
      Top             =   12120
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   10200
      TabIndex        =   10
      Top             =   12120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD"
      Height          =   495
      Left            =   7320
      TabIndex        =   9
      Top             =   12120
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      DataField       =   "occupation"
      DataSource      =   "Data1"
      Height          =   975
      Left            =   4680
      TabIndex        =   8
      Text            =   "Enter occupation"
      Top             =   3000
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      DataField       =   "name"
      DataSource      =   "Data1"
      Height          =   855
      Left            =   4680
      TabIndex        =   7
      Text            =   "Enter name"
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Image ImageBetty 
      Height          =   3975
      Left            =   11760
      Picture         =   "form82.frx":AE6E
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image ImageRiya 
      Height          =   3975
      Left            =   11760
      Picture         =   "form82.frx":CFC3
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image ImageSumit 
      Height          =   3975
      Left            =   11760
      Picture         =   "form82.frx":159B5
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image ImageShreya 
      Height          =   3975
      Left            =   11760
      Picture         =   "form82.frx":1BAEC
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
      Left            =   480
      TabIndex        =   6
      Top             =   9960
      Width           =   3135
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
      Width           =   3135
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
      Left            =   480
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
      Left            =   480
      TabIndex        =   3
      Top             =   4560
      Width           =   3015
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
      Left            =   480
      TabIndex        =   2
      Top             =   3120
      Width           =   3015
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
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   3015
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





Private Sub Form_Load()
Set ImageCombo1.ImageList = ImageList1
Set ImageCombo2.ImageList = ImageList2
Set ImageCombo3.ImageList = ImageList3


ImageCombo1.ComboItems.Add , , "Sunday", 1

ImageCombo1.ComboItems.Add , , "Monday", 1

ImageCombo1.ComboItems.Add , , "Tuesday", 1

ImageCombo1.ComboItems.Add , , "Wednesday", 1

ImageCombo1.ComboItems.Add , , "Thursday", 1

ImageCombo1.ComboItems.Add , , "Friday", 1

ImageCombo1.ComboItems.Add , , "Saturday", 1



ImageCombo2.ComboItems.Add , , "January", 1

ImageCombo2.ComboItems.Add , , "February", 1

ImageCombo2.ComboItems.Add , , "March", 1

ImageCombo2.ComboItems.Add , , "April", 1

ImageCombo2.ComboItems.Add , , "May", 1

ImageCombo2.ComboItems.Add , , "June", 1

ImageCombo2.ComboItems.Add , , "July", 1

ImageCombo2.ComboItems.Add , , "August", 1

ImageCombo2.ComboItems.Add , , "September", 1

ImageCombo2.ComboItems.Add , , "October", 1

ImageCombo2.ComboItems.Add , , "November", 1

ImageCombo2.ComboItems.Add , , "December", 1



ImageCombo3.ComboItems.Add , , "1997", 1

ImageCombo3.ComboItems.Add , , "1998", 1

ImageCombo3.ComboItems.Add , , "1999", 1

ImageCombo3.ComboItems.Add , , "2000", 1

ImageCombo3.ComboItems.Add , , "2001", 1

ImageCombo3.ComboItems.Add , , "2002", 1

ImageCombo3.ComboItems.Add , , "2003", 1

ImageCombo3.ComboItems.Add , , "2004", 1

ImageCombo3.ComboItems.Add , , "2005", 1

ImageCombo3.ComboItems.Add , , "2006", 1

ImageCombo3.ComboItems.Add , , "2007", 1

ImageCombo3.ComboItems.Add , , "2008", 1

ImageCombo3.ComboItems.Add , , "2009", 1

ImageCombo3.ComboItems.Add , , "2010", 1

ImageCombo3.ComboItems.Add , , "2011", 1

ImageCombo3.ComboItems.Add , , "2012", 1

ImageCombo3.ComboItems.Add , , "2013", 1

ImageCombo3.ComboItems.Add , , "2014", 1

ImageCombo3.ComboItems.Add , , "2015", 1

ImageCombo3.ComboItems.Add , , "2016", 1

ImageCombo3.ComboItems.Add , , "2017", 1

ImageCombo3.ComboItems.Add , , "2018", 1

ImageCombo3.ComboItems.Add , , "2019", 1

ImageCombo3.ComboItems.Add , , "2020", 1

End Sub

Private Sub ImageCombo1_Click()
Text4.Text = ImageCombo1.Text



End Sub

Private Sub ImageCombo2_Click()
Text5.Text = ImageCombo2.Text

End Sub


Private Sub ImageCombo3_Click()
Text6.Text = ImageCombo3.Text

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

