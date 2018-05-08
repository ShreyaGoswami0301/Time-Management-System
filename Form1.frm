VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "STUDENT LOGIN PAGE"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16440
   FillColor       =   &H00004040&
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   16440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "CANCEL"
      Height          =   855
      Left            =   6120
      MaskColor       =   &H00400040&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008080&
      Caption         =   "LOGIN"
      Height          =   855
      Left            =   3240
      MaskColor       =   &H00008080&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      IMEMode         =   3  'DISABLE
      Left            =   6120
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1920
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   3
      Top             =   840
      Width           =   5655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "LOG INTO YOUR ACCOUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   1275
      Left            =   3960
      TabIndex        =   0
      Top             =   0
      Width           =   7605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim user As String
Dim pass As String
user = "admin"
pass = "admin"
If (user = Text1.Text And pass = Text2.Text) Then
MsgBox "Congratulation! Your login has been successful."
Form2.Show

Else
MsgBox "Sorry! Your login has not been successful."

End If




End Sub

Private Sub Command2_Click()
End

End Sub
