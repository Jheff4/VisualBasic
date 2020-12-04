VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Fraud Audit Trail"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   12570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "\Intrusion Detection System for Online Banking\Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Record"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Empty Database"
      Height          =   495
      Left            =   8640
      TabIndex        =   2
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   10560
      TabIndex        =   1
      Top             =   4560
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Ffrmaudit.frx":0000
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7435
      _Version        =   393216
      BackColorBkg    =   16777215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Data1.Refresh
Do Until Data1.Recordset.EOF
Data1.Recordset.Delete
Data1.Recordset.MoveNext
Loop
MsgBox "Database is emptied successfully...", vbApplicationModal + vbInformation, "Alert"
End
End Sub
