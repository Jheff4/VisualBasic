VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmprofile 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Intrusion Detection System"
   ClientHeight    =   9525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   12435
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   7200
      TabIndex        =   29
      Top             =   4920
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1800
      TabIndex        =   28
      Top             =   5880
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   122880001
      CurrentDate     =   43836
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmprofile.frx":0000
      Left            =   1800
      List            =   "frmprofile.frx":000A
      TabIndex        =   27
      Text            =   "Savings"
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "\Intrusion Detection System for Online Banking\Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Record"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   10920
      TabIndex        =   26
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   495
      Left            =   9240
      TabIndex        =   25
      Top             =   9000
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   24
      Top             =   8400
      Width           =   2415
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   23
      Top             =   7920
      Width           =   2415
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   22
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   21
      Top             =   6360
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1800
      MaxLength       =   16
      TabIndex        =   20
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   19
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   18
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   17
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   16
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   15
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   10680
      Picture         =   "frmprofile.frx":001F
      Stretch         =   -1  'True
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Initial Deposit: =N="
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5760
      TabIndex        =   30
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   12120
      Y1              =   8880
      Y2              =   8880
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   7920
      Width           =   840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   12120
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Log In Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   7440
      Width           =   2040
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PIN:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   6840
      Width           =   315
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CVV:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   6360
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Date:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   5400
      Width           =   1050
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ATM Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   1050
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   12120
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ATM Card Number:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   1365
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   1245
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other Name:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surname:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   675
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
      Caption         =   "Bank Customer Data Profile Registration:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3495
   End
End
Attribute VB_Name = "frmprofile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Data1.Recordset!surname = Text1
Data1.Recordset!firstname = Text2
Data1.Recordset!othername = Text3
Data1.Recordset!account_name = Text4
Data1.Recordset!account_number = Text5
Data1.Recordset!atm_number = Text7
Data1.Recordset!account_type = Combo1
Data1.Recordset!expiry_date = DTPicker1
Data1.Recordset!cvv = Text10
Data1.Recordset!pin = Text11
Data1.Recordset!UserName = Text12
Data1.Recordset!Password = Text13
Data1.Recordset!deposit = Text6

Data1.Recordset.Update
MsgBox "Customers record submitted successfully...", vbApplicationModal + vbInformation, "Alert"
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

End Sub
