VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00004040&
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   11760
   FillColor       =   &H00004080&
   ForeColor       =   &H80000013&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CLICK MENU TO SELECT MATRIX CALCULATOR"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   4080
      TabIndex        =   1
      Top             =   3600
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DETERMINANT OF A  MATRIX  CALCULATOR"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   465
      Left            =   1440
      TabIndex        =   0
      Top             =   1680
      Width           =   8730
   End
   Begin VB.Menu mnumenu 
      Caption         =   "Menu"
      Begin VB.Menu mnu2 
         Caption         =   "2 x 2  Matrix"
      End
      Begin VB.Menu mnu3 
         Caption         =   "3 x 3  Matrix"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text5 = Val(Text1) * Val(Text4) - Val(Text2) * Val(Text3)
End Sub

Private Sub Command2_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""

End Sub

Private Sub mnu2_Click()
Form3.Show , Form1

End Sub

Private Sub mnu3_Click()
Form2.Show , Form1

End Sub
