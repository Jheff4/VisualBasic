VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmenquiry 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Intrusion Detection System"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   11790
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "\Intrusion Detection System for Online Banking\Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Audit"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "\Intrusion Detection System for Online Banking\Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Record"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   375
      Left            =   10080
      TabIndex        =   16
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Process"
      Height          =   375
      Left            =   8640
      TabIndex        =   15
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Submit"
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   10200
      TabIndex        =   12
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   5
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   4
      Top             =   3120
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmenquiry.frx":0000
      Left            =   1920
      List            =   "frmenquiry.frx":000A
      TabIndex        =   3
      Text            =   "Savings"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1920
      MaxLength       =   16
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   122880001
      CurrentDate     =   43836
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   9720
      Picture         =   "frmenquiry.frx":001F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER ACCOUNT BALANCE ENQUIRY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   3795
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Withdrawal:          =N="
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   1590
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   11760
      X2              =   120
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   1050
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Date:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CVV:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   360
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PIN:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Card Number: "
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1020
   End
End
Attribute VB_Name = "frmenquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
frmlogin.Data1.Recordset.Edit
frmlogin.Data1.Recordset!deposit = Val(frmlogin.Data1.Recordset!deposit) - Val(Text6)
frmlogin.Data1.Recordset.Update
Text2 = frmlogin.Data1.Recordset!deposit
MsgBox "Withdrawal successfull. Your account balance is =N= " + Text2, vbApplicationModal + vbInformation, "Alert"

End Sub

Private Sub Command3_Click()
'Data1.Refresh
'Do Until Data1.Recordset.EOF
'If Data1.Recordset!UserName = Text3 And Data1.Recordset!Password = Text4 Then

If frmlogin.Data1.Recordset!atm_number = Val(Text1) And frmlogin.Data1.Recordset!account_type = Combo1 And frmlogin.Data1.Recordset!cvv = Text10 And frmlogin.Data1.Recordset!pin = Text11 Then
MsgBox "Your current balance is =N= " + Text2, vbApplicationModal + vbInformation, "Alert"

Else
MsgBox "Fraud Suspected...", vbApplicationModal + vbInformation, "Alert"

Data2.Recordset.AddNew
Data2.Recordset!transaction = "Balance enquiry attempted"
Data2.Recordset!Date = Date
Data2.Recordset!Time = Time()
Data2.Recordset.Update
Unload Me
End If

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()

End Sub
