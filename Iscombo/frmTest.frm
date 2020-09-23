VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmtest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ISCombo Test"
   ClientHeight    =   7665
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Auto Complete and more information"
      Height          =   2925
      Left            =   210
      TabIndex        =   26
      Top             =   4710
      Width           =   8955
      Begin VB.CommandButton Command7 
         Caption         =   "sort"
         Height          =   315
         Left            =   8010
         TabIndex        =   32
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Remove"
         Height          =   315
         Left            =   6750
         TabIndex        =   30
         Top             =   840
         Width           =   945
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add"
         Height          =   315
         Left            =   5760
         TabIndex        =   29
         Top             =   840
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3780
         TabIndex        =   28
         Top             =   840
         Width           =   1875
      End
      Begin Project1.ISCombo ISCombo11 
         Height          =   315
         Left            =   180
         TabIndex        =   27
         Top             =   840
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "There are items with the text"
         Height          =   195
         Left            =   210
         TabIndex        =   38
         Top             =   2130
         Width           =   1995
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "There are items with the text"
         Height          =   195
         Left            =   210
         TabIndex        =   37
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "There are items with the text"
         Height          =   195
         Left            =   210
         TabIndex        =   36
         Top             =   2670
         Width           =   1995
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "There are items with the text"
         Height          =   195
         Left            =   210
         TabIndex        =   35
         Top             =   1890
         Width           =   1995
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "There are items with the text"
         Height          =   195
         Left            =   210
         TabIndex        =   34
         Top             =   1620
         Width           =   1995
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "There are items with the text"
         Height          =   195
         Left            =   210
         TabIndex        =   33
         Top             =   1350
         Width           =   1995
      End
      Begin VB.Label Label6 
         Caption         =   $"frmTest.frx":0000
         Height          =   465
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   8085
      End
   End
   Begin Project1.ISCombo ISCombo9 
      Height          =   315
      Left            =   4800
      TabIndex        =   24
      Top             =   3810
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   556
      Caption         =   "This one is locked just try to type something here"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      AutoComplete    =   0   'False
   End
   Begin VB.Frame Frame3 
      Caption         =   "Changing the picture of item"
      Height          =   1635
      Left            =   5640
      TabIndex        =   18
      Top             =   2130
      Width           =   3645
      Begin Project1.ISCombo ISCombo6 
         Height          =   345
         Left            =   150
         TabIndex        =   20
         Top             =   300
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   609
         Caption         =   "Press the button for an other picture"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoComplete    =   0   'False
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Next picture"
         Height          =   345
         Left            =   210
         TabIndex        =   19
         Top             =   1170
         Width           =   1545
      End
      Begin Project1.ISCombo ISCombo8 
         Height          =   345
         Left            =   150
         TabIndex        =   22
         Top             =   690
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   609
         Caption         =   "Press the button for an other picture"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoComplete    =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Changing text of item"
      Height          =   1905
      Left            =   5610
      TabIndex        =   15
      Top             =   150
      Width           =   3675
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2970
         Top             =   270
      End
      Begin Project1.ISCombo ISCombo5 
         Height          =   375
         Left            =   150
         TabIndex        =   17
         Top             =   930
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Caption         =   "For the current time click here"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlign       =   1
         AutoComplete    =   0   'False
      End
      Begin Project1.ISCombo ISCombo7 
         Height          =   375
         Left            =   150
         TabIndex        =   21
         Top             =   1380
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Caption         =   "For the current time click here"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlign       =   2
         AutoComplete    =   0   'False
      End
      Begin VB.Label Label5 
         Caption         =   $"frmTest.frx":00E9
         Height          =   645
         Left            =   60
         TabIndex        =   16
         Top             =   240
         Width           =   3585
      End
   End
   Begin VB.Frame ShowImage 
      Caption         =   "Showing Images"
      Height          =   2145
      Left            =   150
      TabIndex        =   6
      Top             =   2130
      Width           =   4515
      Begin VB.CommandButton Command4 
         Caption         =   "Remove Last"
         Height          =   405
         Left            =   3210
         TabIndex        =   23
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add Picture"
         Height          =   375
         Left            =   1710
         TabIndex        =   14
         Top             =   1560
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add Text"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   1395
      End
      Begin Project1.ISCombo ISCombo2 
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   609
         Caption         =   "Iscombo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoComplete    =   0   'False
      End
      Begin Project1.ISCombo ISCombo3 
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   609
         Caption         =   "Iscombo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Showpictures    =   1
         AutoComplete    =   0   'False
      End
      Begin Project1.ISCombo ISCombo4 
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   609
         Caption         =   "Iscombo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Showpictures    =   2
         AutoComplete    =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "2-Never"
         Height          =   195
         Left            =   3800
         TabIndex        =   12
         Top             =   1170
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "1-Always"
         Height          =   195
         Left            =   3800
         TabIndex        =   11
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "0-Auto"
         Height          =   195
         Left            =   3800
         TabIndex        =   10
         Top             =   270
         Width           =   465
      End
   End
   Begin Project1.ISCombo ISCombo1 
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   1530
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   609
      Caption         =   "Please give a rating"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTest.frx":0178
      AutoComplete    =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select a Flag"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      Begin Project1.ISCombo iccountry 
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   330
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   609
         Caption         =   "Select a flag"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmTest.frx":05CA
         AutoComplete    =   0   'False
      End
   End
   Begin VB.PictureBox pSelected 
      Height          =   1335
      Left            =   3660
      ScaleHeight     =   1275
      ScaleWidth      =   1830
      TabIndex        =   1
      Top             =   120
      Width           =   1890
      Begin VB.Image imgFlag 
         Height          =   1320
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1800
      End
   End
   Begin MSComctlLib.ImageList ilFlags 
      Left            =   1320
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0A1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":12C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1718
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2414
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2868
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":3110
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":3564
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":39B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":3E0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4260
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":46B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":53B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":5804
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":5C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":60AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6500
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6954
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6DA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":71FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":7650
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":7AA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":7C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":7D5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Project1.ISCombo ISCombo10 
      Height          =   315
      Left            =   4800
      TabIndex        =   25
      Top             =   4200
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   556
      Caption         =   "This one is not locked you can just type here"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoComplete    =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "How do you rate this Control??"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3495
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''      ProjectName:    ISCombo Test.
''      Author:         Alfredo Córdova Pérez ( fred_cpp )
''      e-mail:         fred_cpp@hotmail.com
''                      fred_cpp@yahoo.com.mx
''
''      Description:
''
''      I've Got a lot of problemas with the VB' combo, I couldn't detect
''      when the user changes the selection, and, those combos are really ugly :(
''      so, I decided made one better.
''      you know, you can use this freely, just give me credit.
''      Votes and suggestions are wellcome.
''

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const STR_LINK          As String = "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=34300&txtForceRefresh=52200205773948"

Private Sub cmdClose_Click()
    Unload Me
    End
End Sub

Private Sub Command1_Click()
ISCombo2.AddItem Time$
ISCombo3.AddItem Time$
ISCombo4.AddItem Time$
End Sub

Private Sub Command2_Click()
ISCombo2.AddItem Time$, , ImageList1.ListImages(1).Picture
ISCombo3.AddItem Time$, , ImageList1.ListImages(1).Picture
ISCombo4.AddItem Time$, , ImageList1.ListImages(1).Picture
End Sub

Private Sub Command3_Click()
Dim x As Integer
x = ilFlags.ListImages.Count - 1
Randomize Timer
x = Rnd * x + 1
Set ISCombo8.Picture = ilFlags.ListImages(x).Picture
ISCombo6.ChangeItemPicture 0, ilFlags.ListImages(x).Picture
End Sub

Private Sub Command4_Click()
ISCombo2.RemoveItemByNumber ISCombo2.ItemCount
ISCombo3.RemoveItemByNumber ISCombo3.ItemCount
ISCombo4.RemoveItemByNumber ISCombo4.ItemCount
End Sub

Private Sub Command5_Click()
ISCombo11.AddItem Text1.text
End Sub

Private Sub Command6_Click()
ISCombo11.RemoveItemByText Text1.text, False
End Sub

Private Sub Command7_Click()
If Command7.Caption = "sort" Then

Command7.Caption = "unsort"

ISCombo11.Sorted = True

Else

Command7.Caption = "sort"

ISCombo11.Sorted = False

End If
End Sub


Private Sub Form_Load()
    '
    Dim ni As Long
    '' Add Some Flags to this Combo
    For ni = 1 To 26
        iccountry.AddItem "Flag #: " & ni, , ilFlags.ListImages(ni).Picture
    Next ni
    
    ISCombo1.AddItem "Poor", , ImageList1.ListImages(1).Picture
    ISCombo1.AddItem "Below Average", , ImageList1.ListImages(1).Picture
    ISCombo1.AddItem "Average", , ImageList1.ListImages(2).Picture
    ISCombo1.AddItem "Good", , ImageList1.ListImages(3).Picture
    ISCombo1.AddItem "Excellent", , ImageList1.ListImages(3).Picture
        
    ISCombo2.AddItem "hallo"
    ISCombo3.AddItem "hallo"
    ISCombo4.AddItem "hallo"
    
    ISCombo5.AddItem Time$

    ISCombo6.AddItem "Press the button for an other picture"
    ISCombo8.AddItem "Press the button for an other picture"

'    ISCombo1.Picture = ImageList1.ListImages(2).Picture
End Sub

'' This event is generated when user clicks au item list
Private Sub icCountry_ItemClick(iItem As Integer)
    imgFlag.Picture = ilFlags.ListImages(iItem).Picture
End Sub

Private Sub Text1_Change()
Dim x As Integer, p As Integer, k As Integer, j As Integer, z As Integer

z = -1
For x = 0 To ISCombo11.ItemCount
z = ISCombo11.SearchNumberByText(Text1.text, z + 1, False, False)
    If z = -1 Then
    Exit For
    Else
    p = p + 1
    End If
Next

Label7.Caption = "There are " + Trim(Str(p)) + " items with the text '" + Text1.text + "' in it. (With SearchNumberByText)"

z = -1
For x = 0 To ISCombo11.ItemCount
z = ISCombo11.SearchNumberByText(Text1.text, z + 1, False, True)
    If z = -1 Then
    Exit For
    Else
    k = k + 1
    End If
Next

If k = -1 Then k = 0
Label8.Caption = "There are " + Trim(Str(k)) + " items with the exact text '" + Text1.text + "'. (With SearchNumberByText)"

z = -1
For x = 0 To ISCombo11.ItemCount
z = ISCombo11.SearchNumberByText(Text1.text, z + 1, True, True)
    If z = -1 Then
    Exit For
    Else
    j = j + 1
    End If
Next

Label9.Caption = "There are " + Trim(Str(j)) + " items with the exact text '" + Text1.text + "' and the exact case. (With SearchNumberByText)"

k = ISCombo11.CountDoubles(Text1.text, , True, True) 'matchcase on
j = ISCombo11.CountDoubles(Text1.text, , False, True) 'matchcase of
z = ISCombo11.CountDoubles(Text1.text, , False, False) 'instr

Label12.Caption = "There are " + Trim(Str(z)) + " items with the text '" + Text1.text + "' in it. (With CountDoubles)"
Label11.Caption = "There are " + Trim(Str(j)) + " items with the text '" + Text1.text + "' in it. (With CountDoubles)"
Label10.Caption = "There are " + Trim(Str(k)) + " items with the text '" + Text1.text + "' in it. (With CountDoubles)"

End Sub

Private Sub Timer1_Timer()
ISCombo5.ChangeItemText 0, Time$
ISCombo7.Caption = Time$
End Sub
