VERSION 5.00
Begin VB.Form frmupdate 
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "SEARCH"
      Height          =   495
      Left            =   9120
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3600
      TabIndex        =   10
      Text            =   "Select"
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\video rental system\DATABASE FOR AFFILIATE PROGRAM.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   11400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "RECORDS FOR AFFILIATE PROGRAM"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   11160
      TabIndex        =   9
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RESET"
      Height          =   495
      Left            =   9120
      TabIndex        =   8
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      Height          =   495
      Left            =   7080
      TabIndex        =   7
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\video rental system\DATABASE FOR AFFILIATE PROGRAM.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   11400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "RECORDS FOR AFFILIATE PROGRAM"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   4680
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   3840
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   3000
      Width           =   4815
   End
   Begin VB.Label Label5 
      Caption         =   "RECORD UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "OCCUPATION OF CONTACT"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "PHONE NUMBER OF CONTACT"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "ADDRESS OF CONTACT"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "NAME OF CONTACT"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "frmupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
MsgBox "Please Select Name Of Contact From The List", vbApplicationModal + vbInformation, "Hello Dear"

End Sub

Private Sub Combo1_Click()
Data1.Refresh
Do Until Data1.Recordset.EOF
If Combo1 = Data1.Recordset!name_of_contact Then
Text2 = Data1.Recordset!address_of_contact
Text3 = Data1.Recordset!phone_number_of_contact
Text4 = Data1.Recordset!occupation_of_contact
End If
Data1.Recordset.MoveNext
Loop
End Sub

Private Sub Command1_Click()
If Text1 = "" Then
MsgBox "Please Enter Name Of Contact", vbApplicationModal + vbInformation, "Alert"
Text1.SetFocus
GoTo 10
End If

If Text2 = "" Then
MsgBox "Please Enter Address Of Contact", vbApplicationModal + vbInformation, "Alert"
Text2.SetFocus
GoTo 10
End If

If Text3 = "" Then
MsgBox "Please Enter Phone Number Of Contact", vbApplicationModal + vbInformation, "Alert"
Text3.SetFocus
GoTo 10
End If

If Text4 = "" Then
MsgBox "Please Enter Occupation Of Contact", vbApplicationModal + vbInformation, "Alert"
Text4.SetFocus
GoTo 10
End If


Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset!name_of_contact = UCase(Text1)
Data1.Recordset!address_of_contact = UCase(Text2)
Data1.Recordset!phone_number_of_contact = Val(Text3)
Data1.Recordset!occupation_of_contact = UCase(Text4)

Data1.Recordset.Update
MsgBox "One Record Submitted Successfully!", vbApplicationModal + vbExclamation, "Alert"
Dim ans
ans = MsgBox("Do You Want To Enter Any Other Record?", vbApplicationModal + vbQuestion + vbYesNo, "Alert")
If ans = vbYes Then
Unload Me
frmmenu.Show
End If
If ans = vbNo Then
Unload Me
End If



10
End Sub

Private Sub Command2_Click()
Text1 = ""
Combo1 = Clear
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Data1.Refresh
Do Until Data1.Recordset.EOF
If Combo1 = Data1.Recordset!name_of_contact Then
Text2 = Data1.Recordset!address_of_contact
Text3 = Data1.Recordset!phone_number_of_contact
Text4 = Data1.Recordset!occupation_of_contact
End If
Data1.Recordset.MoveNext
Loop

End Sub

Private Sub Form_Activate()
Combo1 = ""
Do Until Data2.Recordset.EOF
Combo1.AddItem (Data2.Recordset!name_of_contact)
Data2.Recordset.MoveNext
Loop

End Sub

