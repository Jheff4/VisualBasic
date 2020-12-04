VERSION 5.00
Begin VB.Form frmmenu 
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   11160
      TabIndex        =   10
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RESET"
      Height          =   495
      Left            =   9120
      TabIndex        =   9
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      Height          =   495
      Left            =   7080
      TabIndex        =   8
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
      Left            =   11160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "RECORDS FOR AFFILIATE PROGRAM"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   4800
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   3960
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   3000
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   3600
      TabIndex        =   1
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Label Label4 
      Caption         =   "OCCUPATION OF CONTACT"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "PHONE NUMBER OF CONTACT"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "ADDRESS OF CONTACT"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "NAME OF CONTACT"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

