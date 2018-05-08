VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   11730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18660
   LinkTopic       =   "Form1"
   ScaleHeight     =   11730
   ScaleWidth      =   18660
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Suvajit\Desktop\project vb 3.0\TIME_DATABASE.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TIME_MANAGE"
      Top             =   10440
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   10320
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   10320
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   10320
      Width           =   3135
   End
   Begin VB.TextBox Text18 
      DataField       =   "1_2"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   14040
      TabIndex        =   36
      Top             =   8880
      Width           =   4335
   End
   Begin VB.TextBox Text17 
      DataField       =   "0_1"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   14040
      TabIndex        =   35
      Top             =   8040
      Width           =   4335
   End
   Begin VB.TextBox Text16 
      DataField       =   "23_0"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   14040
      TabIndex        =   34
      Top             =   7200
      Width           =   4335
   End
   Begin VB.TextBox Text15 
      DataField       =   "22_23"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   14040
      TabIndex        =   33
      Top             =   6240
      Width           =   4335
   End
   Begin VB.TextBox Text14 
      DataField       =   "21_22"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   14040
      TabIndex        =   32
      Top             =   5280
      Width           =   4335
   End
   Begin VB.TextBox Text13 
      DataField       =   "20_21"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   14040
      TabIndex        =   31
      Top             =   4200
      Width           =   4335
   End
   Begin VB.TextBox Text12 
      DataField       =   "18_19"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   14040
      TabIndex        =   30
      Top             =   3120
      Width           =   4335
   End
   Begin VB.TextBox Text11 
      DataField       =   "17_18"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   14040
      TabIndex        =   29
      Top             =   2040
      Width           =   4335
   End
   Begin VB.TextBox Text10 
      DataField       =   "16_17"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   14040
      TabIndex        =   28
      Top             =   1080
      Width           =   4335
   End
   Begin VB.TextBox Text9 
      DataField       =   "15_16"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   4080
      TabIndex        =   18
      Top             =   8880
      Width           =   4455
   End
   Begin VB.TextBox Text8 
      DataField       =   "14_15"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   4080
      TabIndex        =   17
      Top             =   8040
      Width           =   4455
   End
   Begin VB.TextBox Text7 
      DataField       =   "13_14"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   4080
      TabIndex        =   16
      Top             =   7200
      Width           =   4455
   End
   Begin VB.TextBox Text6 
      DataField       =   "12_13"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   4080
      TabIndex        =   15
      Top             =   6360
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      DataField       =   "11_12"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   4080
      TabIndex        =   14
      Top             =   5400
      Width           =   4455
   End
   Begin VB.TextBox Text4 
      DataField       =   "10_11"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   4080
      TabIndex        =   13
      Top             =   4320
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      DataField       =   "9_10"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   4080
      TabIndex        =   12
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      DataField       =   "8_9"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   4080
      TabIndex        =   11
      Top             =   2160
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "7_8"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   4080
      TabIndex        =   10
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "1:00-2:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10320
      TabIndex        =   27
      Top             =   8880
      Width           =   3135
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "00:00-1:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10320
      TabIndex        =   26
      Top             =   8040
      Width           =   3135
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "23:00-00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10320
      TabIndex        =   25
      Top             =   7200
      Width           =   3135
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "22:00-23:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10320
      TabIndex        =   24
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "21:00-22:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10320
      TabIndex        =   23
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "20:00-21:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10320
      TabIndex        =   22
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "18:00-19:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10320
      TabIndex        =   21
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "17:00-18:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10320
      TabIndex        =   20
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "16:00-17:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10320
      TabIndex        =   19
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MANAGE YOUR TIME HOURLY"
      BeginProperty Font 
         Name            =   "Myriad Set OT Md"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   240
      Width           =   15615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "15:00-16:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   9000
      Width           =   2655
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "14:00-15:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "13:00-14:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "12:00-13:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "11:00-12:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "10:00-11:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "9:00-10:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "8:00-9:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "7:00-8:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew

End Sub

Private Sub Command2_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "
Text9.Text = " "
Text10.Text = " "
Text11.Text = " "
Text12.Text = " "
Text13.Text = " "
Text14.Text = " "
Text15.Text = " "
Text16.Text = " "
Text17.Text = " "
Text18.Text = " "


End Sub

Private Sub Command3_Click()
Data1.Recordset.Update
Data1.Recordset.Requery
Data1.Refresh

End Sub

Private Sub Command4_Click()
Data1.Recordset.AddNew


End Sub

Private Sub Command5_Click()
Data1.Recordset.Update


End Sub
