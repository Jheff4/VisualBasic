VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmentry 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   12
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   11
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\LOAN MANAGEMENT SYSTEM\database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "RECORD"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   3480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   -2147483633
      Format          =   108920833
      CurrentDate     =   44140
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000F&
      Height          =   405
      Left            =   2040
      TabIndex        =   8
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      Height          =   405
      Left            =   2040
      TabIndex        =   6
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   405
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   405
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000018&
      BorderWidth     =   5
      X1              =   240
      X2              =   11760
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "NET PAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "STAFF ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000018&
      BorderWidth     =   5
      X1              =   240
      X2              =   11760
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "STAFF NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LOAN ENTRY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Then
MsgBox "Please enter staff name", vbApplicationModal + vbExclamation, "Alert"
Text1.SetFocus
GoTo 10
End If

If Text2 = "" Then
MsgBox "Please enter ID", vbApplicationModal + vbExclamation, "Alert"
Text2.SetFocus
GoTo 10
End If

If Text3 = "" Then
MsgBox "Please enter department", vbApplicationModal + vbExclamation, "Alert"
Text3.SetFocus
GoTo 10
End If

If Text4 = "" Then
MsgBox "Please enter net pay", vbApplicationModal + vbExclamation, "Alert"
Text4.SetFocus
GoTo 10
End If



Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset!STAFF_NAME = UCase(Text1)
Data1.Recordset!STAFF_ID = Text2
Data1.Recordset!DEPARTMENT = UCase(Text3)
Data1.Recordset!Date = DTPicker1
Data1.Recordset!NET_PAY = Val(Text4)
Data1.Recordset.Update

MsgBox "Record successfully submitted", vbApplicationModal + vbInformation, "Alert"
Dim ans
ans = MsgBox("Do you want to enter any other record", vbApplicationModal + vbQuestion + vbYesNo, "Alert")
If ans = vbYes Then
Text1 = ""
Text2 = ""
Text3 = ""
DTPicker1 = Date
Text4 = ""

'Unload Me
'frmentry.Show
Else
Unload Me
End If
10
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
