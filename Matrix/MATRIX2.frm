VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form2 
   BackColor       =   &H00000040&
   Caption         =   "Form2"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13245
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   13245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   975
      Left            =   9240
      TabIndex        =   10
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox e 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   855
      Left            =   3000
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "3 X 3  MATRIX"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   4695
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   5415
      Begin VB.TextBox i 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   855
         Left            =   3960
         TabIndex        =   9
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox h 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   855
         Left            =   2160
         TabIndex        =   8
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox g 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   855
         Left            =   360
         TabIndex        =   7
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox f 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   855
         Left            =   3960
         TabIndex        =   6
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox d 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   855
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox c 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   855
         Left            =   3960
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox b 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   855
         Left            =   2160
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox a 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000003&
         Height          =   855
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   855
      Left            =   8640
      TabIndex        =   13
      Top             =   1440
      Width           =   2895
      ForeColor       =   64
      BackColor       =   -2147483646
      Caption         =   "CLEAR"
      Size            =   "5106;1508"
      FontName        =   "MV Boli"
      FontEffects     =   1073741831
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   855
      Left            =   8640
      TabIndex        =   12
      Top             =   3120
      Width           =   2895
      ForeColor       =   64
      BackColor       =   -2147483646
      Caption         =   "EVALUATE"
      Size            =   "5106;1508"
      FontName        =   "MV Boli"
      FontEffects     =   1073741831
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OUTPUT"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   615
      Left            =   7440
      TabIndex        =   11
      Top             =   5040
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
End Sub

Private Sub CommandButton1_Click()
If a = "" Or b = "" Or c = "" Or d = "" Then
Text10 = 0
Else

Text10 = (Val(a) * ((e * i) - (f * h))) - (Val(b) * ((d * i) - (f * g))) + (Val(c) * ((d * h) - (e * g)))
End If
End Sub

Private Sub CommandButton2_Click()
a = ""
b = ""
c = ""
d = ""
e = ""
f = ""
g = ""
h = ""
i = ""
Text10 = ""

End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub Text10_Click()
'MsgBox "Output Cannot Be Editted", vbApplicationModal + vbExclamation, "Alert"
'Text10 = ""
End Sub
