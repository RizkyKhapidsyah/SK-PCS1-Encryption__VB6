VERSION 5.00
Begin VB.Form FrmEncrypt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Encrypt Test Area"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Pub and N"
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generate Key"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Text            =   "0"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Text            =   "0"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Text            =   "0"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   5775
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   5880
      X2              =   120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label10 
      Caption         =   "Public:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Private:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "N:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "Text to Encrypt:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "Encrypted Text/ Text to Decrypt:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "FrmEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FrmMain.Text3 = Text6
FrmMain.Text4 = Text8

End Sub

Private Sub Command2_Click()
GenKey 256, 1000
Text6 = Pub
Text7 = Prv
Text8 = n
PubI = 0
PrvI = 0
End Sub

Private Sub Command3_Click()
If Text6 = "0" Then GoTo nokey
If Text8 = "0" Then GoTo nokey
On Error Resume Next

Pub = Text6 ' set keys
n = Text8

Text10 = EncryptBk(Text9) ' encrypt data

Exit Sub
nokey:
MsgBox "You Must have the Keys - Public and N to Encrypt the data", vbCritical, "No Keys"
End Sub

Private Sub Command4_Click()
If Text7 = "0" Then GoTo nokey
If Text8 = "0" Then GoTo nokey
On Error Resume Next

Prv = Text7 ' set keys
n = Text8

Text9 = DecryptBk(Text10) ' decrypt data

Exit Sub
nokey:
MsgBox "You Must have the Keys - Private and N to Encrypt the data", vbCritical, "No Keys"
End Sub
