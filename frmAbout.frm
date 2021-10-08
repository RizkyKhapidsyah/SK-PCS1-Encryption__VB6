VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   4320
      Width           =   975
   End
   Begin VB.Line Line3 
      X1              =   5880
      X2              =   120
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label3 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "How To..."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5880
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5880
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "ABOUT PCS-1 Private Key Finder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label9 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload frmAbout
End Sub
