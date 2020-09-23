VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   345
      Left            =   3540
      TabIndex        =   4
      Top             =   3000
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   1275
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "aboutiscombop.frx":0000
      Top             =   1200
      Width           =   4485
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   0
      Y1              =   3390
      Y2              =   0
   End
   Begin VB.Line Line3 
      X1              =   4650
      X2              =   4650
      Y1              =   3390
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   -30
      X2              =   4650
      Y1              =   3390
      Y2              =   3390
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label3 
      Caption         =   "You are free to use it and improve it, only if you just give us our credit. After all we worked hard for it. Enjoy!"
      Height          =   585
      Left            =   60
      TabIndex        =   3
      Top             =   2640
      Width           =   4485
   End
   Begin VB.Label Label2 
      Caption         =   "The best Image Combo ever. Has lots of features other Image Combo's can only dream of. "
      Height          =   525
      Left            =   30
      TabIndex        =   1
      Top             =   660
      Width           =   4515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ISCombo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   1995
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

