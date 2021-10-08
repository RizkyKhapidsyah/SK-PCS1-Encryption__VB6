VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "PCS Private Key Finder"
   ClientHeight    =   2145
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4920
      TabIndex        =   7
      Text            =   "0"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4920
      TabIndex        =   6
      Text            =   "0"
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Text            =   "10000"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "1"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Password Test"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Result:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   855
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   5760
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      X1              =   3960
      X2              =   3960
      Y1              =   360
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5760
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label6 
      Caption         =   "Scan Range:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Private:"
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "N:"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Public:"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Current:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.Menu goenc 
      Caption         =   "Encrypt"
   End
   Begin VB.Menu abt 
      Caption         =   "About"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub abt_Click()
frmAbout.Show

End Sub

Private Sub Command1_Click()

If Text3 = "0" Then GoTo nokeys
If Text4 = "0" Then GoTo nokeys





Dim enctext As String
PubI = 0
PrvI = 0

Pub = Text3 ' set public key
n = Text4 ' set n key
enctext = EncryptBk("TESTDATA") ' encrypt test data block

MsgBox enctext, vbInformation, "Test String"
t1 = Timer
For i = Text1 To Text2 ' start key test
    Label2.Caption = i
    Pub = Trim(Text3)
    n = Trim(Text4)
    Prv = i ' try private key
On Error Resume Next
    testblock = DecryptBk(enctext) ' decrypt it
    If testblock = "TESTDATA" Then GoTo found ' if it decrypts ok then found it
    DoEvents
Next
    MsgBox "Key Not Found", vbExclamation, "NotFound"
Exit Sub
found:
    t2 = Timer
    MsgBox "Private Key found in " & Format(t2 - t1, "##.####") & " seconds", vbInformation, "Found"
    Text5 = i
    Label9 = i
Exit Sub
nokeys:
MsgBox "You must have Public and N keys to Test", vbCritical, "No Keys"
  
End Sub


Private Sub goenc_Click()
FrmEncrypt.Show

End Sub
