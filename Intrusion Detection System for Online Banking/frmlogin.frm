VERSION 5.00
Begin VB.Form frmlogin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "User Login"
   ClientHeight    =   2760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4800
      TabIndex        =   6
      Text            =   "3"
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "\Intrusion Detection System for Online Banking\Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Record"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "#"
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   4440
      Picture         =   "frmlogin.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   6000
      X2              =   0
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   840
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Refresh
Do Until Data1.Recordset.EOF
If Data1.Recordset!UserName = Text1 And Data1.Recordset!Password = Text2 Then
'frmdeposit.Text3 = Text1
'frmdeposit.Text4 = Text2
frmmenu.Show
GoTo 10
End If
Data1.Recordset.MoveNext
Loop
Text3 = Val(Text3) - 1
MsgBox "Invalid Password, you have just" + " " + Text3 + " " + "login trails left, input the correct password or exit", vbApplicationModal + vbInformation, "Alert"

If Val(Text3) = 0 Then
MsgBox "Invalid username and password. Access denied...", vbApplicationModal + vbInformation, "Alert"
End
End If
10
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()

End Sub
