VERSION 5.00
Begin VB.Form frmstartup 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(New User)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10680
      TabIndex        =   5
      Top             =   240
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      Height          =   195
      Left            =   12600
      TabIndex        =   4
      Top             =   240
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sign Up"
      Height          =   195
      Left            =   11640
      TabIndex        =   3
      Top             =   240
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Intrusion Detection System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5640
      TabIndex        =   2
      Top             =   4080
      Width           =   4665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNITED BANK FOR AFRICA PLC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   1440
      TabIndex        =   0
      Top             =   2760
      Width           =   7260
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   9855
   End
   Begin VB.Image Image1 
      Height          =   2025
      Left            =   1200
      Picture         =   "frmstartup.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   3615
   End
End
Attribute VB_Name = "frmstartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Label4_Click()
frmprofile.Show
End Sub

Private Sub Label5_Click()
frmlogin.Show
End Sub
