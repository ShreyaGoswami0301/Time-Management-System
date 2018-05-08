VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "time"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13845
   LinkTopic       =   "Form2"
   ScaleHeight     =   7785
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   10200
      Top             =   2520
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Kameleon"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   5775
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   10695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Form3.Show

End Sub

Private Sub Timer1_Timer()
Label1.Caption = Time

End Sub
