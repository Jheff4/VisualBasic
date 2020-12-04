VERSION 5.00
Begin VB.Form frmmenu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Intrusion Detection System"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   12270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Fraud Audit Trail"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Balance Enquiry"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Withdrawal"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deposit"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   10320
      Picture         =   "frmmenu.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   12120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Customer Transaction Menu"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmdeposit.Show
End Sub

Private Sub Command2_Click()
frmwithdrawal.Show
End Sub

Private Sub Command3_Click()
frmenquiry.Text2 = frmlogin.Data1.Recordset!deposit
frmenquiry.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
frmaudit.Show
End Sub

Private Sub Form_Load()

End Sub
